import sys
import os
from array import array
from datetime import datetime
try:
    import win32com.client
    import pythoncom
except Exception:
    win32com = None
    pythoncom = None

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QLineEdit, QComboBox,
                             QPushButton, QGroupBox, QGridLayout, QScrollArea,
                             QMessageBox, QDoubleSpinBox, QSpinBox, QCheckBox)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QIcon

APP_VERSION = "v0.2.0-ups"
ICON_FILENAME = "plategen_icon.png"

PF_DEFAULT = 0.8
STYLE_REG = 'Consolas'
STYLE_BOLD = 'ConsolasBold'

def ensure_consolas_style(doc):
    """
    Returns a text style named 'ConsolasStyle'.
    NOTE: AutoCAD COM cannot set fonts via Python; style must already exist
    inside the DWG template if you want correct fonts.
    """
    styles = doc.TextStyles
    try:
        st = styles.Item("ConsolasStyle")
        return st
    except:
        st = styles.Add("ConsolasStyle")
        return st

def ensure_consolas_bold_style(doc):
    """
    Returns a bold text style named 'ConsolasBold'.
    Font must already be bold inside DWG; COM cannot set it.
    """
    styles = doc.TextStyles
    try:
        st = styles.Item("ConsolasBold")
        return st
    except:
        st = styles.Add("ConsolasBold")
        return st

def make_safearray_3d(points):
    arr = array('d')
    for x, y, z in points:
        arr.extend([float(x), float(y), float(z)])
    if win32com is None:
        return arr
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, arr)

def make_point_variant(x, y, z=0.0):
    if win32com is None:
        return (float(x), float(y), float(z))
    arr = array('d', [float(x), float(y), float(z)])
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, arr)

def add_rect(ms, x1, y1, x2, y2):
    pts = [(x1,y1,0),(x2,y1,0),(x2,y2,0),(x1,y2,0),(x1,y1,0)]
    v = make_safearray_3d(pts)
    if win32com is None:
        return None
    pl = ms.AddPolyline(v)
    pl.Closed = True
    return pl

def add_line(ms, x1, y1, x2, y2):
    if win32com is None:
        return None
    p1 = make_point_variant(x1, y1, 0)
    p2 = make_point_variant(x2, y2, 0)
    return ms.AddLine(p1,p2)

def add_text(ms, text, x, y, height, style_name=None):
    if win32com is None:
        return None
    p = make_point_variant(x, y, 0)
    t = ms.AddText(str(text), p, float(height))
    if style_name:
        try:
            t.StyleName = style_name
        except Exception:
            pass
    return t

def add_mtext(ms, text, x, y, width, height, style_name=None):
    if win32com is None:
        return None
    p = make_point_variant(x, y, 0)
    mt = ms.AddMText(p, float(width), str(text))
    mt.Height = float(height)
    if style_name:
        try:
            mt.StyleName = style_name
        except Exception:
            pass
    try:
        mt.Attachment = 2
    except Exception:
        pass
    return mt

def insert_scaled_block(ms, block_path, x, y, target_w, target_h):
    if win32com is None:
        return None
    ins_pt = make_point_variant(x, y, 0)
    blk = ms.InsertBlock(ins_pt, block_path, 1.0, 1.0, 1.0, 0)
    try:
        blk.Update()
    except Exception:
        pass
    try:
        bb = blk.GetBoundingBox()
        (xmin, ymin, zmin) = bb[0]
        (xmax, ymax, zmax) = bb[1]
    except Exception:
        return blk
    bw = xmax - xmin
    bh = ymax - ymin
    if bw == 0 or bh == 0:
        return blk
    sx = target_w / bw
    sy = target_h / bh
    s = min(sx, sy)
    blk.XScaleFactor = s
    blk.YScaleFactor = s
    blk.ZScaleFactor = s
    blk.Update()
    try:
        bb2 = blk.GetBoundingBox()
        (xmin2, ymin2, zmin2) = bb2[0]
    except Exception:
        return blk
    dx = x - xmin2
    dy = y - ymin2
    blk.Move(make_point_variant(0, 0, 0), make_point_variant(dx, dy, 0))    
    # Try to explode and remove the block reference so any embedded raster
    # or block content becomes native entities. This helps avoid the issue
    # where only the last inserted block's image displays until objects are
    # moved or the drawing is regenerated.
    deleted = False
    try:
        # get document from ModelSpace if available
        doc = None
        try:
            doc = ms.Parent
        except Exception:
            doc = None

        try:
            # Explode the block instance into its constituent entities
            blk.Explode()
            deleted = True
        except Exception:
            deleted = False

        try:
            # delete the original block reference if possible
            if deleted:
                blk.Delete()
        except Exception:
            pass

        try:
            # Force a regen to refresh display
            if doc is not None:
                try:
                    doc.SendCommand("_REGEN ")
                except Exception:
                    try:
                        # fallback: call Regen method if available
                        doc.Regen(0)
                    except Exception:
                        pass
        except Exception:
            pass
    except Exception:
        pass

    # return None if we removed the block reference, otherwise return the block
    return None if deleted else blk

def draw_rating_plate_ups(doc, config):
    """Draw a simple UPS rating plate. Computed rated power = kVA * PF_DEFAULT."""
    plate_w = config.get('plate_width', 150.0)
    plate_h = config.get('plate_height', 125.0)
    margin = config.get('margin', 3.0)
    ox = config.get('offset_x', 100.0)
    oy = config.get('offset_y', 100.0)

    ms = doc.ModelSpace

    # Ensure Consolas styles exist and use their names
    try:
        style_reg = ensure_consolas_style(doc)
        style_bold = ensure_consolas_bold_style(doc)
        style_reg_name = style_reg.Name
        style_bold_name = style_bold.Name
    except Exception:
        style_reg_name = STYLE_REG
        style_bold_name = STYLE_BOLD

    outer_top = oy + plate_h
    outer_bottom = oy

    # Frames
    add_rect(ms, ox, outer_bottom, ox + plate_w, outer_top)
    add_rect(ms, ox + margin, outer_bottom + margin, ox + plate_w - margin, outer_top - margin)

    ux1 = ox + margin
    ux2 = ox + plate_w - margin
    uy2 = outer_top - margin
    y = uy2

    # Optional: draw simple dimension lines (horizontal below and vertical left)
    if config.get('show_dimensions', False):
        try:
            dt = float(config.get('dim_text_height', 3.0))
        except Exception:
            dt = 3.0
        ext = 8.0
        arrow = 1.6
        # gap between plate edge and dimension extension (in mm)
        gap = float(config.get('dim_gap', 3.0))
        # Use outer rectangle edges so dimensions cover full plate
        outer_left = ox
        outer_right = ox + plate_w

        # horizontal dimension below plate (use outer edges)
        dim_y = outer_bottom - ext
        # extension lines from outer edges down to dimension line, leaving a small gap
        add_line(ms, outer_left, outer_bottom - gap, outer_left, dim_y)
        add_line(ms, outer_right, outer_bottom - gap, outer_right, dim_y)
        # main dim line across full outer width
        add_line(ms, outer_left, dim_y, outer_right, dim_y)
        # simple arrowheads at ends
        add_line(ms, outer_left, dim_y, outer_left + arrow, dim_y + 1)
        add_line(ms, outer_left, dim_y, outer_left + arrow, dim_y - 1)
        add_line(ms, outer_right, dim_y, outer_right - arrow, dim_y + 1)
        add_line(ms, outer_right, dim_y, outer_right - arrow, dim_y - 1)
        wtext = config.get('dim_width_override') or f"{plate_w:g} mm"
        # center text under the dimension line
        add_text(ms, wtext, (outer_left + outer_right) / 2 - 8, dim_y - 6, dt, style_name=style_reg_name)

        # vertical dimension left of plate (use outer edges for references)
        dim_x = outer_left - ext
        # extension lines from outer left to the vertical dimension line, leaving a small gap
        add_line(ms, outer_left - gap, outer_bottom, dim_x, outer_bottom)
        add_line(ms, outer_left - gap, outer_top, dim_x, outer_top)
        # main vertical dim line
        add_line(ms, dim_x, outer_bottom, dim_x, outer_top)
        # arrowheads for vertical dim (small horizontal ticks)
        add_line(ms, dim_x, outer_top, dim_x + 1, outer_top - arrow)
        add_line(ms, dim_x, outer_top, dim_x - 1, outer_top - arrow)
        add_line(ms, dim_x, outer_bottom, dim_x + 1, outer_bottom + arrow)
        add_line(ms, dim_x, outer_bottom, dim_x - 1, outer_bottom + arrow)
        htext = config.get('dim_height_override') or f"{plate_h:g} mm"
        # place height text to the left of the vertical dim line
        import math

        txt = add_text(ms, htext, dim_x - 5,
                    (outer_bottom + outer_top) / 2 - 2,
                    dt,
                    style_name=style_reg_name)

        # Rotate 90 degrees
        txt.Rotation = math.radians(90)
    row_h = 12.0
    text_h = 4.0
    # Smaller text height for unequal frequency variations
    text_h_small = 3.5

    # PRODUCT row
    y_bottom = y - 10
    param_offset_right = 60
    param_offset_top = 7
    add_rect(ms, ux1, y_bottom, ux2, y)
    # Add vertical line separator in PRODUCT row
    add_line(ms, ux1 + param_offset_right - 3, y_bottom, ux1 + param_offset_right - 3, y)
    product_text = config.get('product_text', 'DEFAULT UPS')
    add_text(ms, 'PRODUCT', ux1 + 3, y - param_offset_top, 4.0, style_name=style_bold_name)
    # Draw product description in bold using font override sequence
    try:
        mtext_w = ux2 - (ux1 + param_offset_right) - 4.0
        add_mtext(ms, r"\fConsolas|b1;" + product_text, ux1 + param_offset_right, y - param_offset_top + 4, mtext_w, 4.2, style_name=style_reg_name)
    except Exception:
        # fallback to plain text if MText fails
        add_text(ms, product_text, ux1 + param_offset_right, y - param_offset_top, 4.0, style_name=style_reg_name)
    y = y_bottom

    # INPUT VOLTAGE - use smaller font if frequency variation is unequal
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    # Add vertical line separator in INPUT VOLTAGE row
    add_line(ms, ux1 + param_offset_right - 3, y_bottom, ux1 + param_offset_right - 3, y)
    add_text(ms, 'INPUT VOLTAGE', ux1 + 3, y - param_offset_top, text_h, style_name=style_reg_name)
    input_voltage_text = config.get('input_voltage', '415V, 3 PHASE, 4 WIRES, 50HZ ±5%')
    # Check if input has unequal frequency variation (contains "to")
    input_text_height = text_h_small if ' to ' in input_voltage_text else text_h
    add_text(ms, input_voltage_text, ux1 + param_offset_right, y - param_offset_top, input_text_height, style_name=style_reg_name)
    y = y_bottom

    # OUTPUT VOLTAGE - use smaller font if frequency variation is unequal
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    # Add vertical line separator in OUTPUT VOLTAGE row
    add_line(ms, ux1 + param_offset_right - 3, y_bottom, ux1 + param_offset_right - 3, y)
    add_text(ms, 'OUTPUT VOLTAGE', ux1 + 3, y - param_offset_top, text_h, style_name=style_reg_name)
    output_voltage_text = config.get('output_voltage', '230V, 1PHASE, 2 WIRES, 50HZ')
    # Check if output has unequal frequency variation (contains "to")
    output_text_height = text_h_small if ' to ' in output_voltage_text else text_h
    add_text(ms, output_voltage_text, ux1 + param_offset_right, y - param_offset_top, output_text_height, style_name=style_reg_name)
    y = y_bottom

    # RATED POWER (compute)
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    # Add vertical line separator in RATED POWER row
    add_line(ms, ux1 + param_offset_right - 3, y_bottom, ux1 + param_offset_right - 3, y)
    add_text(ms, 'RATED POWER', ux1 + 3, y - param_offset_top, text_h, style_name=style_reg_name)
    kva = float(config.get('apparent_kva', 0.0))
    pf = float(config.get('pf', PF_DEFAULT))
    rated_kw = round(kva * pf, 3)
    rated_text = f"{rated_kw:g} kW (at {pf:g} PF)"
    add_text(ms, rated_text, ux1 + param_offset_right, y - param_offset_top, text_h, style_name=style_reg_name)
    y = y_bottom

    # SL NO
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    # Add vertical line separator in SL NO row
    add_line(ms, ux1 + param_offset_right - 3, y_bottom, ux1 + param_offset_right - 3, y)
    add_text(ms, 'SL. NO.', ux1 + 3, y - param_offset_top, text_h, style_name=style_reg_name)
    serial = config.get('serial', '')
    add_text(ms, serial, ux1 + param_offset_right, y - param_offset_top, text_h, style_name=style_reg_name)
    y = y_bottom

    # YEAR
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    # Add vertical line separator in YEAR row
    add_line(ms, ux1 + param_offset_right - 3, y_bottom, ux1 + param_offset_right - 3, y)
    add_text(ms, 'YEAR OF MFG.', ux1 + 3, y - param_offset_top, text_h, style_name=style_reg_name)
    add_text(ms, str(config.get('year', datetime.now().year)), ux1 + param_offset_right, y - param_offset_top, text_h, style_name=style_reg_name)
    y = y_bottom

    # Footer
    FOOT_H = 22
    y_footer_top = y
    y_footer_bottom = y_footer_top - FOOT_H
    add_line(ms, ux1, y_footer_top, ux2, y_footer_top)
    # Footer: title (bold) and address/contact lines
    # Footer title in bold using font override
    try:
        add_mtext(ms, r"\fConsolas|b1;LIVELINE ELECTRONICS", ux1 + 3, y_footer_top - 4, 200, 4.2, style_name=style_reg_name)
    except Exception:
        add_mtext(ms, 'LIVELINE ELECTRONICS', ux1 + 3, y_footer_top - 3, 200, 4.0, style_name=style_bold_name)
    # Address and contact
    addr_y = y_footer_top - 14
    add_text(ms, 'North Ramchandrapur, Narendrapur, Kolkata : 700103, WB', ux1 + 3, addr_y, 2.6, style_name=style_reg_name)
    add_text(ms, 'Telefax : 033 2477 2094', ux1 + 3, addr_y - 5, 2.6, style_name=style_reg_name)
    add_text(ms, 'Email : info@livelineindia.com', ux1 + 3, addr_y - 10, 2.6, style_name=style_reg_name)

    # Logo
    logo_w = 45
    logo_h = 40
    logo_block = os.path.abspath("liveline_logo.dwg")
    insert_scaled_block(ms, logo_block, ux2 - 45.5, y_footer_bottom -6.5, logo_w, logo_h)

    try:
        doc.SendCommand("_ZOOM _E ")
    except Exception:
        pass

    print("UPS rating plate generated.")


class UPSRatingPlateGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('UPS Rating Plate Generator')
        self.setMinimumSize(750, 600)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        title = QLabel('UPS Rating Plate Generator')
        title_font = QFont('Consolas', 14)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scrollw = QWidget()
        scroll.setWidget(scrollw)
        v = QVBoxLayout(scrollw)

        v.addWidget(self.create_general_group())
        v.addWidget(self.create_voltage_group())
        v.addWidget(self.create_power_group())
        v.addWidget(self.create_dimensions_group())
        v.addWidget(self.create_serial_group())

        v.addStretch()
        layout.addWidget(scroll)

        btn = QPushButton('Generate Rating Plate')
        btn.clicked.connect(self.generate_plate)
        layout.addWidget(btn)

        self.update_rated_power()
        self.update_voltage_display()
        # Initialize wire config enable/disable state
        self.on_input_voltage_changed()
        self.on_output_voltage_changed()

    def create_general_group(self):
        g = QGroupBox('General')
        l = QGridLayout()
        l.addWidget(QLabel('Project Number (Job):'), 0, 0)
        self.project_no = QSpinBox(); self.project_no.setRange(1,9999); self.project_no.setValue(1000)
        l.addWidget(self.project_no, 0, 1)
        l.addWidget(QLabel('Order Number (OP):'), 0, 2)
        self.order_no = QSpinBox(); self.order_no.setRange(1,9999); self.order_no.setValue(1)
        l.addWidget(self.order_no, 0, 3)
        l.addWidget(QLabel('Year:'), 1, 0)
        self.year = QSpinBox(); self.year.setRange(2000,2100); self.year.setValue(datetime.now().year)
        l.addWidget(self.year, 1, 1)
        g.setLayout(l)
        return g

    def create_voltage_group(self):
        g = QGroupBox('Voltage Configuration')
        l = QGridLayout()

        # Input Voltage
        l.addWidget(QLabel('Input Voltage:'), 0, 0)
        self.input_voltage_combo = QComboBox()
        self.input_voltage_combo.addItems(['230V', '415V'])
        self.input_voltage_combo.setCurrentText('415V')
        self.input_voltage_combo.currentTextChanged.connect(self.on_input_voltage_changed)
        l.addWidget(self.input_voltage_combo, 0, 1)

        self.input_wires_label = QLabel('Input Wire Config (for 415V):')
        l.addWidget(self.input_wires_label, 0, 2)
        self.input_wires_combo = QComboBox()
        self.input_wires_combo.addItems(['3 WIRES', '4 WIRES'])
        self.input_wires_combo.setCurrentText('4 WIRES')
        self.input_wires_combo.currentTextChanged.connect(self.update_voltage_display)
        l.addWidget(self.input_wires_combo, 0, 3)

        # Output Voltage
        l.addWidget(QLabel('Output Voltage:'), 1, 0)
        self.output_voltage_combo = QComboBox()
        self.output_voltage_combo.addItems(['230V', '415V'])
        self.output_voltage_combo.setCurrentText('230V')
        self.output_voltage_combo.currentTextChanged.connect(self.on_output_voltage_changed)
        l.addWidget(self.output_voltage_combo, 1, 1)

        self.output_wires_label = QLabel('Output Wire Config (for 415V):')
        l.addWidget(self.output_wires_label, 1, 2)
        self.output_wires_combo = QComboBox()
        self.output_wires_combo.addItems(['3 WIRES', '4 WIRES'])
        self.output_wires_combo.setCurrentText('4 WIRES')
        self.output_wires_combo.currentTextChanged.connect(self.update_voltage_display)
        l.addWidget(self.output_wires_combo, 1, 3)

        # Frequency Configuration
        l.addWidget(QLabel('Base Frequency (Hz):'), 2, 0)
        self.base_freq = QDoubleSpinBox()
        self.base_freq.setRange(40, 70)
        self.base_freq.setValue(50)
        self.base_freq.setSingleStep(0.1)
        self.base_freq.valueChanged.connect(self.update_voltage_display)
        l.addWidget(self.base_freq, 2, 1)

        l.addWidget(QLabel('Input Freq Var Up (%):'), 3, 0)
        self.input_freq_up = QDoubleSpinBox()
        self.input_freq_up.setRange(0, 50)
        self.input_freq_up.setValue(5)
        self.input_freq_up.setSingleStep(0.1)
        self.input_freq_up.valueChanged.connect(self.update_voltage_display)
        l.addWidget(self.input_freq_up, 3, 1)

        l.addWidget(QLabel('Input Freq Var Down (%):'), 3, 2)
        self.input_freq_down = QDoubleSpinBox()
        self.input_freq_down.setRange(0, 50)
        self.input_freq_down.setValue(5)
        self.input_freq_down.setSingleStep(0.1)
        self.input_freq_down.valueChanged.connect(self.update_voltage_display)
        l.addWidget(self.input_freq_down, 3, 3)

        # Output frequency variation checkbox
        l.addWidget(QLabel('Show Output Freq Var:'), 4, 0)
        self.show_output_freq_var = QCheckBox()
        self.show_output_freq_var.setChecked(False)
        self.show_output_freq_var.stateChanged.connect(self.update_voltage_display)
        l.addWidget(self.show_output_freq_var, 4, 1)

        l.addWidget(QLabel('Output Freq Var Up (%):'), 5, 0)
        self.output_freq_up = QDoubleSpinBox()
        self.output_freq_up.setRange(0, 50)
        self.output_freq_up.setValue(1)
        self.output_freq_up.setSingleStep(0.1)
        self.output_freq_up.valueChanged.connect(self.update_voltage_display)
        l.addWidget(self.output_freq_up, 5, 1)

        l.addWidget(QLabel('Output Freq Var Down (%):'), 5, 2)
        self.output_freq_down = QDoubleSpinBox()
        self.output_freq_down.setRange(0, 50)
        self.output_freq_down.setValue(1)
        self.output_freq_down.setSingleStep(0.1)
        self.output_freq_down.valueChanged.connect(self.update_voltage_display)
        l.addWidget(self.output_freq_down, 5, 3)

        # Display fields
        l.addWidget(QLabel('Generated Input:'), 6, 0)
        self.input_voltage_display = QLineEdit()
        self.input_voltage_display.setReadOnly(True)
        l.addWidget(self.input_voltage_display, 6, 1, 1, 3)

        l.addWidget(QLabel('Generated Output:'), 7, 0)
        self.output_voltage_display = QLineEdit()
        self.output_voltage_display.setReadOnly(True)
        l.addWidget(self.output_voltage_display, 7, 1, 1, 3)

        g.setLayout(l)
        return g

    def create_power_group(self):
        g = QGroupBox('UPS Data')
        l = QGridLayout()

        l.addWidget(QLabel('Product suffix:'), 0, 0)
        self.product_suffix = QLineEdit('UPS-1 PANEL')
        l.addWidget(self.product_suffix, 0, 1)
        # Number of UPS units (1..10). If >1, a BYP (bypass) plate will also be generated
        l.addWidget(QLabel('Unit Count:'), 0, 2)
        self.unit_count = QSpinBox()
        self.unit_count.setRange(1, 10)
        self.unit_count.setValue(1)
        l.addWidget(self.unit_count, 0, 3)

        l.addWidget(QLabel('Apparent Power (kVA):'), 1, 0)
        self.kva = QDoubleSpinBox(); self.kva.setRange(0.1, 10000.0); self.kva.setValue(7.5); self.kva.setSingleStep(0.1)
        self.kva.valueChanged.connect(self.update_rated_power)
        l.addWidget(self.kva, 1, 1)

        l.addWidget(QLabel('Power Factor:'), 1, 2)
        self.pf = QDoubleSpinBox(); self.pf.setRange(0.1, 1.0); self.pf.setSingleStep(0.01); self.pf.setValue(PF_DEFAULT)
        self.pf.valueChanged.connect(self.update_rated_power)
        l.addWidget(self.pf, 1, 3)

        l.addWidget(QLabel('Rated Power (kW):'), 2, 0)
        self.rated_display = QLineEdit(); self.rated_display.setReadOnly(True)
        l.addWidget(self.rated_display, 2, 1)

        g.setLayout(l)
        return g

    def create_dimensions_group(self):
        g = QGroupBox('Dimensions')
        l = QGridLayout()

        l.addWidget(QLabel('Show Dimensions:'), 0, 0)
        self.show_dimensions = QCheckBox()
        self.show_dimensions.setChecked(True)
        l.addWidget(self.show_dimensions, 0, 1)

        l.addWidget(QLabel('Dimension text height:'), 0, 2)
        self.dim_text_height = QDoubleSpinBox()
        self.dim_text_height.setRange(0.5, 20.0)
        self.dim_text_height.setValue(3.0)
        self.dim_text_height.setSingleStep(0.1)
        l.addWidget(self.dim_text_height, 0, 3)

        # Width override
        l.addWidget(QLabel('Override Width Text:'), 1, 0)
        self.dim_width_override_chk = QCheckBox('Use override')
        l.addWidget(self.dim_width_override_chk, 1, 1)
        self.dim_width_override = QLineEdit('')
        self.dim_width_override.setPlaceholderText('e.g. 185 mm')
        self.dim_width_override.setEnabled(False)
        l.addWidget(self.dim_width_override, 1, 2, 1, 2)
        self.dim_width_override_chk.stateChanged.connect(lambda s: self.dim_width_override.setEnabled(self.dim_width_override_chk.isChecked()))

        # Height override
        l.addWidget(QLabel('Override Height Text:'), 2, 0)
        self.dim_height_override_chk = QCheckBox('Use override')
        l.addWidget(self.dim_height_override_chk, 2, 1)
        self.dim_height_override = QLineEdit('')
        self.dim_height_override.setPlaceholderText('e.g. 105 mm')
        self.dim_height_override.setEnabled(False)
        l.addWidget(self.dim_height_override, 2, 2, 1, 2)
        self.dim_height_override_chk.stateChanged.connect(lambda s: self.dim_height_override.setEnabled(self.dim_height_override_chk.isChecked()))

        g.setLayout(l)
        return g

    def create_serial_group(self):
        g = QGroupBox('Serial and Identifiers')
        l = QGridLayout()
        # Reuse Project/Order from General settings; only need serial suffix here
        l.addWidget(QLabel('Serial suffix:'), 0, 0)
        self.sn_suffix = QLineEdit('UPS')
        l.addWidget(self.sn_suffix, 0, 1)
        g.setLayout(l)
        return g

    def update_rated_power(self):
        kva = float(self.kva.value())
        pf = float(self.pf.value())
        rated = kva * pf
        self.rated_display.setText(f"{rated:g} kW (at {pf:g} PF)")

    def on_input_voltage_changed(self):
        """Enable/disable input wire config based on voltage selection."""
        is_415 = self.input_voltage_combo.currentText() == '415V'
        self.input_wires_combo.setEnabled(is_415)
        self.input_wires_label.setEnabled(is_415)
        self.update_voltage_display()

    def on_output_voltage_changed(self):
        """Enable/disable output wire config based on voltage selection."""
        is_415 = self.output_voltage_combo.currentText() == '415V'
        self.output_wires_combo.setEnabled(is_415)
        self.output_wires_label.setEnabled(is_415)
        self.update_voltage_display()

    def format_voltage_string(self, voltage, wires, freq_base, freq_up, freq_down, show_freq_var=True):
        """Format voltage string with proper phase, wires, and frequency variation."""
        if voltage == '230V':
            phase_str = '1PH, 2 WIRES'
        else:  # 415V
            phase_str = f'3PH, {wires}'
        
        # Format frequency variation
        freq_str = f'{freq_base:g}HZ'
        
        if show_freq_var:
            if abs(freq_up - freq_down) < 0.01:  # Equal variation
                freq_str += f' ±{freq_up:g}%'
            else:  # Unequal variation
                freq_str += f' +{freq_up:g}% to -{freq_down:g}%'
        
        return f'{voltage}, {phase_str}, {freq_str}'

    def update_voltage_display(self):
        """Update voltage display fields based on current selections."""
        # Input voltage
        input_v = self.input_voltage_combo.currentText()
        input_w = self.input_wires_combo.currentText()
        freq = self.base_freq.value()
        in_up = self.input_freq_up.value()
        in_down = self.input_freq_down.value()
        
        input_str = self.format_voltage_string(input_v, input_w, freq, in_up, in_down, True)
        self.input_voltage_display.setText(input_str)
        
        # Output voltage
        output_v = self.output_voltage_combo.currentText()
        output_w = self.output_wires_combo.currentText()
        out_up = self.output_freq_up.value()
        out_down = self.output_freq_down.value()
        show_out_var = self.show_output_freq_var.isChecked()
        
        output_str = self.format_voltage_string(output_v, output_w, freq, out_up, out_down, show_out_var)
        self.output_voltage_display.setText(output_str)

    def get_config(self):
        cfg = {}
        cfg['plate_width'] = 185.0
        cfg['plate_height'] = 105.0
        cfg['offset_x'] = 100.0
        cfg['offset_y'] = 100.0
        cfg['product_text'] = f"{self.kva.value():g}kVA {self.product_suffix.text()}"
        cfg['apparent_kva'] = self.kva.value()
        cfg['pf'] = self.pf.value()
        cfg['unit_count'] = int(self.unit_count.value())
        cfg['input_voltage'] = self.input_voltage_display.text()
        cfg['output_voltage'] = self.output_voltage_display.text()
        # Dimensions config
        cfg['show_dimensions'] = bool(getattr(self, 'show_dimensions', False) and self.show_dimensions.isChecked())
        cfg['dim_text_height'] = float(getattr(self, 'dim_text_height', 3.0).value() if hasattr(self, 'dim_text_height') else 3.0)
        if getattr(self, 'dim_width_override_chk', None) and self.dim_width_override_chk.isChecked():
            cfg['dim_width_override'] = self.dim_width_override.text()
        else:
            cfg['dim_width_override'] = ''
        if getattr(self, 'dim_height_override_chk', None) and self.dim_height_override_chk.isChecked():
            cfg['dim_height_override'] = self.dim_height_override.text()
        else:
            cfg['dim_height_override'] = ''
        year = int(self.year.value())
        cfg['year'] = year
        yy1 = year % 100
        yy2 = (year + 1) % 100
        cfg['serial'] = f"LL/{yy1:02d}-{yy2:02d}/{self.project_no.value()}-OP{self.order_no.value()}/{self.sn_suffix.text()}"
        return cfg

    def generate_plate(self):
        base_cfg = self.get_config()

        # Build list of configurations to generate
        to_generate = []
        unit_count = base_cfg.get('unit_count', 1)
        kva = base_cfg.get('apparent_kva', 0.0)

        for i in range(1, unit_count + 1):
            cfg = dict(base_cfg)
            cfg['product_text'] = f"{kva:g}kVA UPS-{i} PANEL"
            yy1 = cfg['year'] % 100
            yy2 = (cfg['year'] + 1) % 100
            cfg['serial'] = f"LL/{yy1:02d}-{yy2:02d}/{self.project_no.value()}-OP{self.order_no.value()}/UPS{i}"
            to_generate.append(cfg)

        # Add bypass plate if more than one UPS unit
        if unit_count > 1:
            cfg = dict(base_cfg)
            cfg['product_text'] = f"{kva:g}kVA BYPASS PANEL"
            yy1 = cfg['year'] % 100
            yy2 = (cfg['year'] + 1) % 100
            cfg['serial'] = f"LL/{yy1:02d}-{yy2:02d}/{self.project_no.value()}-OP{self.order_no.value()}/BYP"
            to_generate.append(cfg)

        # If AutoCAD not present, show planned plates
        if win32com is None:
            QMessageBox.information(self, 'Planned Plates', 'AutoCAD not available. The following plates would be generated:\n\n' + '\n'.join([g['product_text'] + '  ->  ' + g['serial'] for g in to_generate]))
            return

        try:
            acad = win32com.client.Dispatch('AutoCAD.Application')
            doc = acad.ActiveDocument
        except Exception as e:
            QMessageBox.critical(self, 'AutoCAD Error', f'Could not access AutoCAD: {e}')
            return

        failures = []
        # Layout plates in a grid: side-by-side (max 2 per row) then stack rows below.
        # spacing between plates (horizontal/vertical)
        spacing = base_cfg.get('inter_plate_spacing', 10.0)
        # extra gap to add to the right of each plate (useful when placing multiple plates)
        extra_right = base_cfg.get('multi_right_gap', 10.0)
        # extra gap to add below each plate
        extra_bottom = base_cfg.get('multi_bottom_gap', 10.0)
        per_row = 2
        plate_w = base_cfg.get('plate_width', 185.0)
        plate_h = base_cfg.get('plate_height', 105.0)

        for idx, cfg in enumerate(to_generate):
            col = idx % per_row
            row = idx // per_row
            # compute offsets: start from base offset and shift right by (plate_w + spacing + extra_right) per column
            cfg_offset_x = base_cfg.get('offset_x', 100.0) + col * (plate_w + spacing + extra_right)
            # shift down by (plate_h + spacing + extra_bottom) per row
            cfg_offset_y = base_cfg.get('offset_y', 100.0) - row * (plate_h + spacing + extra_bottom)
            cfg['offset_x'] = cfg_offset_x
            cfg['offset_y'] = cfg_offset_y
            try:
                draw_rating_plate_ups(doc, cfg)
            except Exception as e:
                failures.append((cfg.get('product_text', '<unknown>'), str(e)))

        if not failures:
            QMessageBox.information(self, 'Done', f'Generated {len(to_generate)} plates in AutoCAD.')
        else:
            msg = 'Some plates failed to generate:\n' + '\n'.join([f'{p}: {err}' for p, err in failures])
            QMessageBox.warning(self, 'Partial Failure', msg)


def main():
    app = QApplication(sys.argv)
    window = UPSRatingPlateGUI()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
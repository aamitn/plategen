# ACDB DCDB RATING PLATE GENERATOR
import sys
import os
import math
from array import array
from math import ceil
from datetime import datetime
try:
    import win32com.client
    import pythoncom
except Exception:
    win32com = None
    pythoncom = None

from PyQt6.QtWidgets import (QApplication, QDialog, QMainWindow, QWidget,
                             QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                             QComboBox, QPushButton, QGroupBox, QGridLayout,
                             QMessageBox, QSpinBox, QDoubleSpinBox,
                             QListWidget, QListWidgetItem)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QIcon

# Application version
try:
    with open("appver.txt", "r") as f:
        base_version = f.read().strip()
except Exception:
    base_version = "v0.0.0"

    
# 2. Append the required suffix
APP_VERSION = f"{base_version}-db"
APP_NAME = 'DB Rating Plate Generator'


def make_safearray_3d(points):
    arr = array('d')
    for x, y, z in points:
        arr.extend([float(x), float(y), float(z)])
    if win32com is None:
        return arr
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, arr)


def make_point_variant(x, y, z=0.0):
    # Accept either numbers or tuples
    if isinstance(x, (list, tuple)) and y is None:
        x_, y_, z_ = x
        x, y, z = x_, y_, z_
    if win32com is None:
        return (float(x), float(y), float(z))
    arr = array('d', [float(x), float(y), float(z)])
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, arr)


def add_rect(ms, x1, y1, x2, y2):
    pts = [(x1, y1, 0), (x2, y1, 0), (x2, y2, 0), (x1, y2, 0), (x1, y1, 0)]
    v = make_safearray_3d(pts)
    if win32com is None:
        # fallback: no-op
        return None
    pl = ms.AddPolyline(v)
    pl.Closed = True
    return pl


def add_line(ms, x1, y1, x2, y2):
    if win32com is None:
        return None
    return ms.AddLine(make_point_variant(x1, y1), make_point_variant(x2, y2))

def get_consolas_style(doc):
    """
    Returns a TextStyle named 'Consolas' if it exists.
    If not found, attempts to create it, but COM cannot apply font properties.
    """
    styles = doc.TextStyles
    try:
        return styles.Item("Consolas")
    except Exception:
        # Create style if missing (font must be preconfigured in DWG)
        try:
            return styles.Add("Consolas")
        except:
            return None


def add_text(ms, text, x, y, height):
    if win32com is None:
        return None

    t = ms.AddText(str(text), make_point_variant(x, y), float(height))

    # Assign Consolas style if available
    try:
        doc = ms.Parent
        st = get_consolas_style(doc)
        if st:
            t.StyleName = "Consolas"
    except:
        pass

    return t


def add_mtext(ms, text, x, y, width, height):
    if win32com is None:
        return None

    mt = ms.AddMText(make_point_variant(x, y), float(width), str(text))
    mt.Height = float(height)

    try:
        mt.Attachment = 2
    except:
        pass

    # Apply Consolas
    try:
        doc = ms.Parent
        st = get_consolas_style(doc)
        if st:
            mt.StyleName = "Consolas"
    except:
        pass

    return mt



def add_dimension_linear(
        ms,
        x1, y1,
        x2, y2,
        dimline_x, dimline_y,
        override_text=None,
        vertical=False,
        arrow_size=3.0
    ):
    if win32com is None:
        mx = (x1 + x2) / 2.0
        my = (y1 + y2) / 2.0
        raw_val = abs(x2 - x1) if abs(x2 - x1) > 0 else abs(y2 - y1)
        val = override_text if override_text else f"{raw_val:.1f}"
        add_text(ms, str(val), mx, my, 6.0)
        return None

    try:
        dim = ms.AddDimAligned(
            make_point_variant(x1, y1),
            make_point_variant(x2, y2),
            make_point_variant(dimline_x, dimline_y)
        )

        # Dimension text height
        try:
            dim.TextHeight = 3.0
        except:
            pass

        # Increase extension line gap
        try:
            dim.ExtLineOffset = 6.0   # change to 10, 20 etc. to increase gap
            dim.Update()
        except:
            pass

        # Arrow size
        try:
            dim.ArrowheadSize = arrow_size
        except:
            pass

        # Dimension text override
        # If user provided override, use it; otherwise force one decimal
        # place precision for displayed measurement (e.g. 150.0).
        try:
            raw_val = abs(x2 - x1) if abs(x2 - x1) > 0 else abs(y2 - y1)
            if override_text:
                try:
                    dim.TextOverride = str(override_text)
                except:
                    pass
            else:
                try:
                    dim.TextOverride = f"{raw_val:.1f}"
                except:
                    pass
        except:
            pass

        # Rotate text vertically if needed
        if vertical:
            import math
            try:
                dim.TextRotation = math.pi / 2
            except:
                pass

        return dim

    except:
        return None




def draw_db_plate(doc, config, suppress_zoom=False):
    """Draw ACDB/DCDB rating plate based on config and add width/height dimensions.
    Uses config keys:
      - plate_width (float)
      - plate_height (float)
      - offset_x, offset_y
      - override_width (string) - show this instead of measured width if non-empty
      - override_height (string)
    """
    plate_w = float(config.get('plate_width', 150.0))
    plate_h = float(config.get('plate_height', 95.0))
    ox = float(config.get('offset_x', 100.0))
    oy = float(config.get('offset_y', 100.0))
    margin = float(config.get('margin', 3.0))

    ms = doc.ModelSpace

    outer_top = oy + plate_h
    outer_bottom = oy

    # inner usable x range
    ux1 = ox + margin
    ux2 = ox + plate_w - margin
    y = outer_top - margin

    row_h = 10.5
    txt_h = 3.2

    param_offset_right = 40
    param_offset_top = 6

    # PRODUCT row
    product_top = y
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, 'PRODUCT', ux1 + 2, y - param_offset_top - 1, txt_h)
    add_mtext(ms, config.get('product_text', 'AC DISTRIBUTION BOARD'), ux1 + param_offset_right, y - param_offset_top + 2.5, ux2 - (ux1 + param_offset_right), txt_h + 0.2)
    y = y_bottom

    # INPUT VOLTAGE row
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, 'INPUT VOLTAGE', ux1 + 2, y - param_offset_top, txt_h-0.4)
    add_text(ms, config.get('input_voltage', ''), ux1 + param_offset_right, y - param_offset_top, txt_h)
    y = y_bottom

    # INCOMER row
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, 'INCOMER', ux1 + 2, y - param_offset_top, txt_h)
    add_text(ms, config.get('incomer', ''), ux1 + param_offset_right, y - param_offset_top, txt_h)
    y = y_bottom

    # OUTGOINGS
    out_list = config.get('outgoings', [])
    n_out = len(out_list)
    per_row = 2
    groups = [out_list[i:i+per_row] for i in range(0, n_out, per_row)] if n_out else [[]]
    for grp in groups:
        y_bottom = y - row_h
        add_rect(ms, ux1, y_bottom, ux2, y)
        add_text(ms, 'OUTGOING', ux1 + 2, y - param_offset_top, txt_h)
        parts = []
        for it in grp:
            rating = it.get('rating', '')
            poles = it.get('poles', '')
            btype = it.get('type', '')
            count = it.get('count', 1)
            parts.append(f"{rating}A {poles}P {btype} - {count} NOS.")
        combined = ' ; '.join(parts) if parts else ''
        try:
            add_mtext(ms, combined, ux1 + param_offset_right, y - param_offset_top + 3, ux2 - (ux1 + param_offset_right) + 24, txt_h - 0.2)
        except Exception:
            add_text(ms, combined, ux1 + param_offset_right, y - param_offset_top, txt_h)
        y = y_bottom

    # SL NO and YEAR
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, 'SL. NO.', ux1 + 2, y - param_offset_top, txt_h)
    add_text(ms, config.get('serial', ''), ux1 + param_offset_right, y - param_offset_top, txt_h)
    y = y_bottom

    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, 'YEAR OF MFG.', ux1 + 2, y - param_offset_top, txt_h)
    add_text(ms, str(config.get('year', datetime.now().year)), ux1 + param_offset_right, y - param_offset_top, txt_h)
    y = y_bottom

    # vertical separator
    try:
        vline_x = ux1 + 37
        vline_top = product_top
        vline_bottom = y + row_h
        add_line(ms, vline_x, vline_top, vline_x, vline_bottom - 10.5)
    except Exception:
        pass

    # Footer + logo
    FOOT_H = 20
    y_footer_top = y
    y_footer_bottom = y_footer_top - FOOT_H
    add_line(ms, ux1, y_footer_top, ux2, y_footer_top)
    try:
        add_mtext(ms, r"\fConsolas|b1;LIVELINE ELECTRONICS", ux1 + 2, y_footer_top - 6, 160, 3.6)
    except Exception:
        add_mtext(ms, 'LIVELINE ELECTRONICS', ux1 + 2, y_footer_top, 160, 3.6)

    try:
        add_text(ms, 'North Ramchandrapur, Narendrapur, Kolkata : 700103', ux1 + 2, y_footer_top - 15, 2.6)
        add_text(ms, 'Telefax : 033 2477 2094', ux1 + 2, y_footer_top - 19, 2.4)
        add_text(ms, 'Email : info@livelineindia.com', ux1 + 2, y_footer_top - 23, 2.4)
    except Exception:
        try:
            add_text(ms, 'North Ramchandrapur, Narendrapur, Kolkata : 700103 | Tel: 033 2477 2094 | info@livelineindia.com', ux1 + 2, y_footer_top - 12, 2.6)
        except Exception:
            pass

    # logo
    logo_block = os.path.abspath('liveline_logo.dwg')
    if os.path.exists(logo_block):
        try:
            ins_pt = make_point_variant(ux2 - 62, y_footer_bottom - 16)
            blk = ms.InsertBlock(ins_pt, logo_block, 38.0, 38.0, 38.0, 0)
            try:
                blk.Update()
            except Exception:
                pass
            try:
                doc.SendCommand('_REGEN ')
            except Exception:
                try:
                    doc.Regen(0)
                except Exception:
                    pass
        except Exception:
            pass

    # Re-draw outer & inner frames (final)
    try:
        new_outer_bottom = min(outer_bottom, y_footer_bottom - margin - 7)
        add_rect(ms, ox, new_outer_bottom, ox + plate_w, outer_top)
        add_rect(ms, ox + margin, new_outer_bottom + margin, ox + plate_w - margin, outer_top - margin)
    except Exception:
        pass

    # ---------------------------
    # Dimensions: width (horizontal) and height (vertical)
    # ---------------------------
    try:
        # user-provided override strings
        override_w = config.get('override_width', '').strip()
        override_h = config.get('override_height', '').strip()

        # If frames were redrawn to extend bottom, prefer that as final bottom
        final_bottom = locals().get('new_outer_bottom', outer_bottom)

        # width dimension: between (ox, final_bottom) and (ox+plate_w, final_bottom)
        # place dim line below the bottom by configured gap (defaults to 12 mm)
        # This controls the height of the short extension lines on both ends
        # of the dimension line (use config key 'dim_gap').
        dim_gap = float(config.get('dim_gap', 12.0))
        # keep previous behavior of respecting margin: place dim line
        # at `margin + dim_gap` away from the plate edge so extension
        # lines are visibly clear of the plate.
        dim_offset = margin + dim_gap
        # dimline point for width: mid-x, lower than bottom
        dim_w_x = ox + plate_w / 2.0
        dim_w_y = final_bottom - dim_offset

        # For AutoCAD AddDimAligned we need a point on the dimension line outside the measured segment.
        add_dimension_linear(
            ms,
            ox, final_bottom,
            ox + plate_w, final_bottom,
            dim_w_x, dim_w_y,
            override_text=override_w if override_w else None,
            arrow_size=5.0
        )

        # height dimension: between (ox, outer_bottom) and (ox, outer_top)
        # place dimline point left of the box by 12 units
        dim_h_x = ox - dim_offset
        # place height dim centered between final bottom and outer_top
        dim_h_y = final_bottom + (outer_top - final_bottom) / 2.0
        add_dimension_linear(
            ms,
            ox, final_bottom,  # P1: Bottom Y extent (use extended bottom)
            ox, outer_top,     # P2: Top Y extent
            dim_h_x, dim_h_y,  # Dimension line location
            override_text=override_h if override_h else None,
            vertical=True,     # Flag for text rotation
            arrow_size=5.0
        )



    except Exception:
        pass

    # zoom extents if possible (can be suppressed when drawing many plates)
    if not suppress_zoom:
        try:
            doc.SendCommand('_ZOOM _E ')
        except Exception:
            try:
                doc.Regen(0)
            except Exception:
                pass


def draw_plates_grid(doc, config):
    """Draw multiple plates in a grid layout based on config keys:
      - units (int): total number of plates to draw
      - cols (int): number of columns per row
      - plate_gap (float): gap between plates in mm

    The function calls `draw_db_plate` repeatedly, adjusting `offset_x` and
    `offset_y` for each tile. Zoom is suppressed for intermediate tiles to
    avoid repeated zoom-extents calls.
    """
    units = int(config.get('units', 1))
    # Auto-compute number of columns to form a near-square layout:
    # cols = ceil(sqrt(units)). Examples: 4 -> 2, 5 -> 3
    cols = int(ceil(math.sqrt(units))) if units >= 1 else 1
    plate_gap = float(config.get('plate_gap', 20.0))

    plate_w = float(config.get('plate_width', 150.0))
    plate_h = float(config.get('plate_height', 95.0))
    base_ox = float(config.get('offset_x', 100.0))
    base_oy = float(config.get('offset_y', 100.0))

    for i in range(units):
        r = i // cols
        c = i % cols
        cfg = dict(config)
        cfg['offset_x'] = base_ox + c * (plate_w + plate_gap)
        # move rows downwards by subtracting in Y (y increases upwards)
        cfg['offset_y'] = base_oy - r * (plate_h + plate_gap)

        # suppress zoom for all but the last plate
        suppress = (i != units - 1)
        draw_db_plate(doc, cfg, suppress_zoom=suppress)


class OutgoingDialog(QDialog):
    def __init__(self, parent=None, data=None):
        super().__init__(parent)
        self.setWindowTitle('Add/Edit Outgoing')
        self.resize(360, 140)
        l = QGridLayout(self)

        l.addWidget(QLabel('Rating (A):'), 0, 0)
        self.rating = QSpinBox(); self.rating.setRange(1, 5000); self.rating.setValue(6)
        l.addWidget(self.rating, 0, 1)

        l.addWidget(QLabel('Poles:'), 0, 2)
        self.poles = QComboBox(); self.poles.addItems(['1','2','3']); self.poles.setCurrentText('2')
        l.addWidget(self.poles, 0, 3)

        l.addWidget(QLabel('Type:'), 1, 0)
        self.btype = QComboBox(); self.btype.addItems(['MCB','MCCB']); l.addWidget(self.btype, 1, 1)

        l.addWidget(QLabel('Count:'), 1, 2)
        self.count = QSpinBox(); self.count.setRange(1, 1000); self.count.setValue(1)
        l.addWidget(self.count, 1, 3)

        btns = QHBoxLayout()
        ok = QPushButton('OK'); ok.clicked.connect(self.accept); btns.addWidget(ok)
        cancel = QPushButton('Cancel'); cancel.clicked.connect(self.reject); btns.addWidget(cancel)
        l.addLayout(btns, 2, 0, 1, 4)

        if data:
            self.rating.setValue(int(data.get('rating', 6)))
            self.poles.setCurrentText(str(data.get('poles', '2')))
            self.btype.setCurrentText(data.get('type', 'MCB'))
            self.count.setValue(int(data.get('count', 1)))

    def get_data(self):
        return {
            'rating': int(self.rating.value()),
            'poles': int(self.poles.currentText()),
            'type': self.btype.currentText(),
            'count': int(self.count.value())
        }


class DBRatingPlateGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(920, 700)

        # Status Bar
        status_bar = self.statusBar()
        version_label = QLabel(f"Version: {APP_VERSION}")
        version_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        version_label.setMinimumWidth(150)
        status_bar.addWidget(version_label, 1) # The stretch factor (1) pushes it to the right

        main = QWidget(); self.setCentralWidget(main)
        L = QVBoxLayout(main)

        title = QLabel(APP_NAME)
        title.setFont(QFont('Consolas', 14))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        L.addWidget(title)

        cfgg = self.create_config_group()
        L.addWidget(cfgg)

        outg = self.create_outgoing_group()
        L.addWidget(outg)

        btn = QPushButton('Generate DB Plate')
        btn.setMinimumHeight(40)
        btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 14px;
                font-weight: bold;
                border-radius: 5px;
                padding: 10px 30px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        btn.clicked.connect(self.generate_plate)
        L.addWidget(btn)

    def create_config_group(self):
        g = QGroupBox('General')
        l = QGridLayout()
        l.addWidget(QLabel('DB Type:'), 0, 0)
        self.db_type = QComboBox(); self.db_type.addItems(['ACDB','DCDB']); l.addWidget(self.db_type, 0, 1)

        l.addWidget(QLabel('Project Number (Job):'), 0, 2)
        self.project_no = QSpinBox(); self.project_no.setRange(1,9999); self.project_no.setValue(1000)
        l.addWidget(self.project_no, 0, 3)

        l.addWidget(QLabel('Order Number (OP):'), 1, 0)
        self.order_no = QSpinBox(); self.order_no.setRange(1,9999); self.order_no.setValue(1)
        l.addWidget(self.order_no, 1, 1)

        l.addWidget(QLabel('Year:'), 1, 2)
        self.year = QSpinBox(); self.year.setRange(2000,2100); self.year.setValue(datetime.now().year)
        l.addWidget(self.year, 1, 3)

        l.addWidget(QLabel('Product suffix:'), 2, 0)
        self.product = QLineEdit(''); self.product.setReadOnly(True)
        l.addWidget(self.product, 2, 1, 1, 3)

        # AC specifics
        l.addWidget(QLabel('AC Voltage:'), 3, 0)
        self.ac_voltage = QComboBox(); self.ac_voltage.addItems(['230V','415V']); l.addWidget(self.ac_voltage, 3, 1)
        l.addWidget(QLabel('AC Wires (for 415V):'), 3, 2)
        self.ac_wires = QComboBox(); self.ac_wires.addItems(['3 WIRES','4 WIRES']); self.ac_wires.setCurrentText('4 WIRES')
        l.addWidget(self.ac_wires, 3, 3)

        l.addWidget(QLabel('Input Voltage:'), 4, 0)
        self.input_voltage = QLineEdit('')
        self.input_voltage.setReadOnly(True)
        l.addWidget(self.input_voltage, 4, 1, 1, 3)

        l.addWidget(QLabel('Incomer:'), 5, 0)
        self.incomer = QLineEdit('80A 3P MCCB')
        l.addWidget(self.incomer, 5, 1, 1, 3)

        # plate geometry controls
        l.addWidget(QLabel('Plate Width (mm):'), 6, 0)
        self.plate_width = QDoubleSpinBox(); self.plate_width.setRange(10.0, 2000.0); self.plate_width.setValue(150.0); self.plate_width.setSingleStep(1.0)
        l.addWidget(self.plate_width, 6, 1)

        l.addWidget(QLabel('Plate Height (mm):'), 6, 2)
        self.plate_height = QDoubleSpinBox(); self.plate_height.setRange(10.0, 2000.0); self.plate_height.setValue(95.0); self.plate_height.setSingleStep(1.0)
        l.addWidget(self.plate_height, 6, 3)

        # override display text for dims
        l.addWidget(QLabel('Override Width Text:'), 7, 0)
        self.override_w = QLineEdit('')
        self.override_w.setPlaceholderText('e.g. 150 mm  OR leave blank for actual')
        l.addWidget(self.override_w, 7, 1)

        l.addWidget(QLabel('Override Height Text:'), 7, 2)
        self.override_h = QLineEdit('')
        self.override_h.setPlaceholderText('e.g. 95 mm  OR leave blank for actual')
        l.addWidget(self.override_h, 7, 3)

        # Dimension gap control: controls distance between plate edge and dim line
        l.addWidget(QLabel('Dimension Gap (mm):'), 8, 0)
        self.dim_gap = QDoubleSpinBox(); self.dim_gap.setRange(0.0, 200.0); self.dim_gap.setValue(3.0); self.dim_gap.setSingleStep(1.0)
        l.addWidget(self.dim_gap, 8, 1)

        # Tiling / duplication controls
        l.addWidget(QLabel('Units (count):'), 9, 0)
        self.units = QSpinBox(); self.units.setRange(1, 1000); self.units.setValue(1)
        l.addWidget(self.units, 9, 1)

        # Columns are auto-computed from Units (no user control)

        l.addWidget(QLabel('Plate Gap (mm):'), 10, 0)
        self.plate_gap = QDoubleSpinBox(); self.plate_gap.setRange(0.0, 500.0); self.plate_gap.setValue(20.0); self.plate_gap.setSingleStep(1.0)
        l.addWidget(self.plate_gap, 10, 1)

        # connect signals
        self.db_type.currentTextChanged.connect(self.on_db_type_changed)
        self.ac_voltage.currentTextChanged.connect(self.update_input_voltage)
        self.ac_wires.currentTextChanged.connect(self.update_input_voltage)

        # initialize UI state
        self.on_db_type_changed(self.db_type.currentText())

        g.setLayout(l)
        return g

    def create_outgoing_group(self):
        g = QGroupBox('Outgoings')
        l = QVBoxLayout()
        self.out_list_widget = QListWidget()
        l.addWidget(self.out_list_widget)

        btns = QHBoxLayout()
        add = QPushButton('Add'); add.clicked.connect(self.add_outgoing); btns.addWidget(add)
        edit = QPushButton('Edit'); edit.clicked.connect(self.edit_outgoing); btns.addWidget(edit)
        remove = QPushButton('Remove'); remove.clicked.connect(self.remove_outgoing); btns.addWidget(remove)
        up = QPushButton('Up'); up.clicked.connect(self.move_up); btns.addWidget(up)
        down = QPushButton('Down'); down.clicked.connect(self.move_down); btns.addWidget(down)
        l.addLayout(btns)

        g.setLayout(l)
        return g

    def on_db_type_changed(self, txt: str):
        """Adjust UI when DB type changes: set product text and input voltage rules."""
        if txt == 'DCDB':
            self.product.setText('DC DISTRIBUTION BOARD')
            # DC: two wires, no phase/frequency
            # Pre-fill with default DC value but allow user edits
            self.input_voltage.setText('110V DC, 2 WIRES')
            self.input_voltage.setReadOnly(False)
            # disable AC-specific controls
            self.ac_voltage.setEnabled(False)
            self.ac_wires.setEnabled(False)
        else:
            self.product.setText('AC DISTRIBUTION BOARD')
            # enable AC controls
            self.ac_voltage.setEnabled(True)
            self.ac_wires.setEnabled(True)
            # For AC, input voltage is auto-generated; keep read-only
            self.input_voltage.setReadOnly(True)
            self.update_input_voltage()

    def update_input_voltage(self):
        v = self.ac_voltage.currentText()
        if v == '230V':
            # single-phase 230V
            self.input_voltage.setText('230V AC, 1 PH, 2 WIRES, 50HZ')
        else:
            # 415V: phase is 3PH, wires from user selection
            wires = self.ac_wires.currentText()
            self.input_voltage.setText(f'415V AC, 3PH, {wires}, 50HZ')

    def add_outgoing(self):
        d = OutgoingDialog(self)
        if d.exec() == QDialog.DialogCode.Accepted:
            data = d.get_data()
            text = f"{data['rating']}A {data['poles']}P {data['type']} - {data['count']} NOS."
            item = QListWidgetItem(text)
            item.setData(Qt.ItemDataRole.UserRole, data)
            self.out_list_widget.addItem(item)

    def edit_outgoing(self):
        itm = self.out_list_widget.currentItem()
        if not itm:
            return
        data = itm.data(Qt.ItemDataRole.UserRole)
        d = OutgoingDialog(self, data)
        if d.exec() == QDialog.DialogCode.Accepted:
            new = d.get_data()
            itm.setText(f"{new['rating']}A {new['poles']}P {new['type']} - {new['count']} NOS.")
            itm.setData(Qt.ItemDataRole.UserRole, new)

    def remove_outgoing(self):
        row = self.out_list_widget.currentRow()
        if row >= 0:
            self.out_list_widget.takeItem(row)

    def move_up(self):
        row = self.out_list_widget.currentRow()
        if row > 0:
            itm = self.out_list_widget.takeItem(row)
            self.out_list_widget.insertItem(row - 1, itm)
            self.out_list_widget.setCurrentRow(row - 1)

    def move_down(self):
        row = self.out_list_widget.currentRow()
        if row < self.out_list_widget.count() - 1 and row >= 0:
            itm = self.out_list_widget.takeItem(row)
            self.out_list_widget.insertItem(row + 1, itm)
            self.out_list_widget.setCurrentRow(row + 1)

    def get_config(self):
        cfg = {}
        cfg['product_text'] = self.product.text()
        cfg['input_voltage'] = self.input_voltage.text()
        cfg['incomer'] = self.incomer.text()
        year = int(self.year.value()) if hasattr(self, 'year') else datetime.now().year
        cfg['year'] = year
        yy1, yy2 = self.compute_fiscal_yy(year)
        proj = int(self.project_no.value()) if hasattr(self, 'project_no') else 0
        op = int(self.order_no.value()) if hasattr(self, 'order_no') else 0
        dtype = self.db_type.currentText() if hasattr(self, 'db_type') else 'ACDB'
        cfg['serial'] = f"LL/{yy1:02d}-{yy2:02d}/{proj}-OP{op}/{dtype}"
        # outgoings
        outs = []
        for i in range(self.out_list_widget.count()):
            itm = self.out_list_widget.item(i)
            outs.append(itm.data(Qt.ItemDataRole.UserRole))
        cfg['outgoings'] = outs
        # plate geometry from UI
        cfg['plate_width'] = float(self.plate_width.value())
        cfg['plate_height'] = float(self.plate_height.value())
        cfg['offset_x'] = 100.0
        cfg['offset_y'] = 100.0
        cfg['margin'] = 3.0
        # dimension gap (distance from plate edge to dimension line)
        cfg['dim_gap'] = float(self.dim_gap.value())
        # tiling / duplication
        cfg['units'] = int(self.units.value()) if hasattr(self, 'units') else 1
        # columns are auto-computed from units; do not take from GUI
        cfg['plate_gap'] = float(self.plate_gap.value()) if hasattr(self, 'plate_gap') else 20.0
        # overrides for dimension text
        cfg['override_width'] = self.override_w.text().strip()
        cfg['override_height'] = self.override_h.text().strip()
        return cfg

    def compute_fiscal_yy(self, year, ref_date=None):
        """Compute two-digit fiscal year start and end for the given year based on a reference date (default: today).

        Assumes fiscal year runs from April (4) to March (3). If the reference month is April or later,
        the fiscal year that includes the given calendar "year" starts in that year (e.g., Apr 2026 -> FY 26-27).
        For Jan-Mar, the fiscal year that contains the given year started in the previous calendar year
        (e.g., Feb 2026 -> FY 25-26).
        """
        if ref_date is None:
            ref_date = datetime.now()
        if ref_date.month >= 4:
            start = year
            end = year + 1
        else:
            start = year - 1
            end = year
        return start % 100, end % 100
    
    def generate_plate(self):
        cfg = self.get_config()
        if win32com is None:
            # No AutoCAD: show planned plate summary and return
            preview = (f"PLANNED PLATE\nProduct: {cfg['product_text']}\n"
                       f"Input Voltage: {cfg['input_voltage']}\n"
                       f"Incomer: {cfg['incomer']}\n"
                       f"Plate WxH: {cfg['plate_width']} x {cfg['plate_height']} mm\n"
                       f"Override W text: {cfg['override_width']}\nOverride H text: {cfg['override_height']}\n"
                       f"Outgoings: {len(cfg['outgoings'])} items")
            QMessageBox.information(self, 'Planned Plate', preview)
            return
        try:
            acad = win32com.client.Dispatch('AutoCAD.Application')
            acad.Visible = True
            template_path = os.path.abspath("acadiso.dwt")
            doc = acad.Documents.Add(template_path)
            # doc = acad.ActiveDocument
        except Exception as e:
            QMessageBox.critical(self, 'AutoCAD Error', f'Could not access AutoCAD: {e}')
            return

        try:
            # Draw plates in grid according to units/cols/plate_gap
            draw_plates_grid(doc, cfg)
            QMessageBox.information(self, 'Done', 'DB plate(s) generated in AutoCAD')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to generate plate(s): {e}')


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    app.setWindowIcon(QIcon.fromTheme("appointment-new"))
    
    w = DBRatingPlateGUI()
    w.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()

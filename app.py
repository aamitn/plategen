import win32com.client
import pythoncom
from array import array
import urllib.request
import json
import webbrowser
import os
import time
import sys
from datetime import datetime
import threading
import re
import subprocess
import base64
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QComboBox, 
                             QPushButton, QGroupBox, QGridLayout, QScrollArea,
                             QMessageBox, QDoubleSpinBox, QSpinBox, QCheckBox)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont, QAction, QIcon

# Application version
APP_VERSION = "v0.0.1"

# Default embedded PNG icon (tiny fallback). Place `plategen_icon.png` next
DEFAULT_ICON_B64 = ("iVBORw0KG")

ICON_FILENAME = "plategen_icon.png"

def ensure_app_icon():
    here = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(here, ICON_FILENAME)
    if not os.path.exists(icon_path):
        try:
            with open(icon_path, 'wb') as f:
                f.write(base64.b64decode(DEFAULT_ICON_B64))
        except Exception:
            return None
    return icon_path

# -----------------------------
# SAFEARRAY helpers
# -----------------------------
def make_safearray_3d(points):
    arr = array('d')
    for x, y, z in points:
        arr.extend([float(x), float(y), float(z)])
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, arr)

def make_point_variant(x, y, z=0.0):
    
    arr = array('d', [float(x), float(y), float(z)])
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, arr)

# -----------------------------
# Text style (Consolas) helper
# -----------------------------
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


def fetch_latest_github_release(repo):
    """Return latest release tag for a given GitHub repo 'owner/repo'."""
    if not repo or '/' not in repo:
        return None, None, 'No repo configured'

    url = f"https://api.github.com/repos/{repo}/releases/latest"
    req = urllib.request.Request(url, headers={"User-Agent": "plategen-agent"})
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            if resp.status != 200:
                return None, None, f"HTTP {resp.status}"
            data = resp.read().decode('utf-8')
            j = json.loads(data)
            tag = j.get('tag_name') or j.get('name')
            html_url = j.get('html_url')
            return tag, html_url, None
    except Exception as e:
        return None, None, str(e)


def compare_versions(local, remote):
    """Compare two version strings. Return -1 if remote>local, 0 if equal, 1 if local>remote."""
    if not local or not remote:
        return 0

    def norm(v):
        # Remove leading 'v' and any non-digit/.- characters
        v2 = v.lstrip('vV')
        v2 = re.sub(r"[^0-9\.\-]", '', v2)
        parts = [int(x) if x.isdigit() else 0 for x in v2.split('.')]
        return parts

    try:
        lv = norm(local)
        rv = norm(remote)
        # Compare element-wise
        for a, b in zip(lv, rv):
            if a < b:
                return -1
            if a > b:
                return 1
        # If all zipped equal, longer list wins
        if len(lv) < len(rv):
            return -1
        if len(lv) > len(rv):
            return 1
        return 0
    except Exception:
        return 0

# -----------------------------
# Primitives: rectangles, lines, text, mtext
# -----------------------------
def add_rect(ms, x1, y1, x2, y2):
    pts = [(x1,y1,0),(x2,y1,0),(x2,y2,0),(x1,y2,0),(x1,y1,0)]
    v = make_safearray_3d(pts)
    pl = ms.AddPolyline(v)
    pl.Closed = True
    return pl

def add_line(ms, x1, y1, x2, y2):
    p1 = make_point_variant(x1, y1, 0)
    p2 = make_point_variant(x2, y2, 0)
    return ms.AddLine(p1,p2)

def add_text(ms, text, x, y, height, style):
    p = make_point_variant(x, y, 0)
    t = ms.AddText(text, p, float(height))
    t.StyleName = style.Name
    return t

def add_mtext(ms, text, x, y, width, height, style):
    p = make_point_variant(x, y, 0)
    mt = ms.AddMText(p, float(width), text)
    mt.Height = float(height)
    mt.StyleName = style.Name
    try:
        mt.Attachment = 2
    except Exception:
        pass
    return mt

def align_label(label, width=8):
    return label.ljust(width)

def add_dimension_aligned(ms, x1, y1, x2, y2, dim_x, dim_y, override_text=None, text_height=None):
    """
    Add an aligned dimension between (x1,y1) and (x2,y2) with
    dimension line passing through (dim_x,dim_y).
    """
    p1 = make_point_variant(x1, y1)
    p2 = make_point_variant(x2, y2)
    p_dim_alligned = make_point_variant(dim_x - 3, dim_y - 3)
    dim = ms.AddDimAligned(p1, p2, p_dim_alligned)

    try:
        dim.TextFill = False
    except Exception:
        pass
    
    dim.TextGap = 1.5

    content = override_text if override_text is not None else "<>"
    final_override = r"\FConsolas;"

    if text_height:
        final_override += r"\H{0};".format(text_height)
    
    final_override += content
    dim.TextOverride = final_override
    
    return dim

def insert_scaled_block(ms, block_path, x, y, target_w, target_h):
    ins_pt = make_point_variant(x, y, 0)
    blk = ms.InsertBlock(ins_pt, block_path, 1.0, 1.0, 1.0, 0)

    try:
        blk.Update()
    except:
        pass

    bb = blk.GetBoundingBox()
    (xmin, ymin, zmin) = bb[0]
    (xmax, ymax, zmax) = bb[1]

    bw = xmax - xmin
    bh = ymax - ymin

    if bw == 0 or bh == 0:
        print("Block has zero geometry.")
        return blk

    sx = target_w / bw
    sy = target_h / bh
    s = min(sx, sy)

    blk.XScaleFactor = s
    blk.YScaleFactor = s
    blk.ZScaleFactor = s

    blk.Update()
    bb2 = blk.GetBoundingBox()
    (xmin2, ymin2, zmin2) = bb2[0]

    dx = x - xmin2
    dy = y - ymin2

    blk.Move(make_point_variant(0, 0, 0), make_point_variant(dx, dy, 0))

    return blk

# -----------------------------
# Main rating plate drawer
# -----------------------------
def draw_rating_plate(doc, config):
    """
    Draw rating plate using configuration from GUI
    """
    plate_w = config.get('plate_width', 150.0)
    plate_h = config.get('plate_height', 100.0)
    margin = config.get('margin', 3.0)
    offset_x = config.get('offset_x', 100.0)
    offset_y = config.get('offset_y', 100.0)
    label_w = config.get('label_w', 40.0)
    product_h = config.get('product_h', 20.0)
    row_h = config.get('row_h', 10.0)
    logo_w = config.get('logo_width', 35.0)
    logo_h = config.get('logo_height', 20.0)
    draw_logo_box = config.get('draw_logo_box', False)
    
    mode = config.get('mode', 'single')
    bottom_extra_single = config.get('bottom_extra_single', 3)
    bottom_extra_dual = config.get('bottom_extra_dual', 13.0)
    bottom_extra_ffcb = config.get('bottom_extra_ffcb', 13.0)
    bottom_extra_dualsf = config.get('bottom_extra_dualsf', 18.0)
    
    dimension_text_size = config.get('dim_text_size', 5)

    style = ensure_consolas_style(doc)
    bold_style = ensure_consolas_bold_style(doc)
    
    ms = doc.ModelSpace

    ox = offset_x
    oy = offset_y
    w = plate_w
    h = plate_h
    
    # Choose bottom extension per mode
    if mode == 'ffcb':
        bottom_extra = float(bottom_extra_ffcb)
    elif mode == 'dual':
        bottom_extra = float(bottom_extra_dual)
    elif mode == 'dualsf':
        bottom_extra = float(bottom_extra_dualsf)
    else:
        bottom_extra = float(bottom_extra_single)

    outer_top = oy + h
    outer_bottom = oy - bottom_extra

    # Outer + inner frames
    add_rect(ms, ox, outer_bottom, ox + w, outer_top)
    add_rect(ms, ox + margin, outer_bottom + margin, ox + w - margin, outer_top - margin)
    # Get dimension override values
    dim_width_override = config.get('dim_width_override', None)
    dim_height_override = config.get('dim_height_override', None)
    
    # Use override if provided, otherwise calculate from actual dimensions
    width_text = f"{dim_width_override:.1f} mm" if dim_width_override else f"{w:.1f} mm"
    height_text = f"{dim_height_override:.1f} mm" if dim_height_override else f"{(h + bottom_extra):.1f} mm"
    
    # Dimensions
    add_dimension_aligned(
        ms, ox, outer_bottom, ox + w, outer_bottom,
        ox + w/2, outer_bottom - 8,
        width_text, text_height=dimension_text_size
    )

    add_dimension_aligned(
        ms, ox, outer_bottom, ox, outer_top,
        ox - 10, outer_top - (h + bottom_extra)/2.0,
        height_text, text_height=dimension_text_size
    )
    
    ux1 = ox + margin
    uy1 = outer_bottom + margin
    ux2 = ox + w - margin
    uy2 = outer_top - margin

    sep_gap = 2.0
    y_top_inner = uy2
    y = y_top_inner

    # PRODUCT row
    y_bottom_product = y - product_h
    add_rect(ms, ux1, y_bottom_product, ux2, y)
    add_text(ms, "PRODUCT", ux1 + 3, y - 12, 4, bold_style)

    vertical_shift = 3.0
    vx = ux1 + label_w + sep_gap + vertical_shift
    data_x = vx + 2.0

    col1_x = data_x
    col1_w = (ux2 - data_x) * 0.5 - 6
    col2_x = data_x + col1_w + 8

    mtext_w = (ux2 - data_x) - 4.0
    mt_top = y - 2
    product_font_h = config.get('product_font_h', 2.0)

    product_desc = config.get('product_desc', 'DEFAULT_PRODUCT_DESCRIPTION')
    add_mtext(ms, r"\fConsolas|b1;" + product_desc, data_x, mt_top, mtext_w, product_font_h, style)

    y = y_bottom_product

    # INPUT VOLTAGE row
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, "INPUT VOLTAGE", ux1 + 3, y - 6, 3.0, style)
    
    input_voltage = config.get('input_voltage', '415V AC, 3 PHASE, 4 WIRES, 50HZ')
    add_text(ms, input_voltage, data_x, y - 6, 3.0, style)
    y = y_bottom

    # 3-column block (OUTPUT VOLTAGE + OUTPUT CURRENT)
    three_top = y

    if mode in ('dual', 'ffcb', 'dualsf'):
        y_header_bottom = y - row_h
        add_rect(ms, ux1, y_header_bottom, ux2, y)
        add_text(ms, "OUTPUT VOLT-AMP", ux1 + 3, y - 6, 2.8, style)

        if mode == 'dual':
            left_label = "CHARGER-I"
            right_label = "CHARGER-II"
            shift_left = -12
        elif mode == 'dualsf':
            left_label = "CHARGER-I"
            right_label = "CHARGER-II"
            shift_left = -18
        else:  # ffcb
            left_label = "FLOAT CHARGER"
            right_label = "FCB CHARGER"
            shift_left = -18

        col1_center_x = col1_x + (col1_w / 2.0)
        col2_w = (ux2 - col2_x) - 4.0
        col2_center_x = col2_x + (col2_w / 2.0)

        add_text(ms, left_label, col1_center_x + shift_left, y - 6, 3.0, bold_style)
        add_text(ms, right_label, col2_center_x + shift_left, y - 6, 3.0, bold_style)

        y = y_header_bottom

    # OUTPUT VOLTAGE
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, "OUTPUT VOLTAGE", ux1 + 3, y - 6, 3.0, style)

    if mode == 'dual':
        float_y = y - 4
        boost_y = y - 8
        ch1_fv = config.get('ch1_float_voltage', 123.75)
        ch1_bv = config.get('ch1_boost_voltage', 126.5)
        ch2_fv = config.get('ch2_float_voltage', 123.75)
        ch2_bv = config.get('ch2_boost_voltage', 126.5)
        
        add_text(ms, f"FLOAT : {ch1_fv}V", col1_x + 2, float_y, 3.0, style)
        add_text(ms, f"BOOST : {ch1_bv}V", col1_x + 2, boost_y, 3.0, style)
        add_text(ms, f"FLOAT : {ch2_fv}V", col2_x + 2, float_y, 3.0, style)
        add_text(ms, f"BOOST : {ch2_bv}V", col2_x + 2, boost_y, 3.0, style)
        
    elif mode == 'dualsf':
        float_y = y - 4
        boost_y = y - 8
        ch1_fv = config.get('ch1_float_voltage', 54.0)
        ch1_bv = config.get('ch1_boost_voltage', 66.0)
        ch2_fv = config.get('ch2_float_voltage', 54.0)
        ch2_bv = config.get('ch2_boost_voltage', 66.0)
        
        add_text(ms, f"FLOAT : {ch1_fv}V", col1_x + 2, float_y, 3.0, style)
        add_text(ms, f"BOOST : {ch1_bv}V", col1_x + 2, boost_y, 3.0, style)
        add_text(ms, f"FLOAT : {ch2_fv}V", col2_x + 2, float_y, 3.0, style)
        add_text(ms, f"BOOST : {ch2_bv}V", col2_x + 2, boost_y, 3.0, style)
        
    elif mode == 'ffcb':
        float_charger_v = config.get('float_charger_voltage', 123.75)
        fcb_fv = config.get('fcb_float_voltage', 123.75)
        fcb_bv = config.get('fcb_boost_voltage', 126.5)
        
        add_text(ms, f"{float_charger_v}V", col1_x + 2, y - 6, 3.0, style)
        float_y = y - 4
        boost_y = y - 8
        add_text(ms, f"FLOAT : {fcb_fv}V", col2_x + 2, float_y, 3.0, style)
        add_text(ms, f"BOOST : {fcb_bv}V", col2_x + 2, boost_y, 3.0, style)
        
    else:  # single
        fv = config.get('float_voltage', 123.75)
        bv = config.get('boost_voltage', 126.5)
        add_text(ms, f"FLOAT : {fv}V", col1_x + 2, y - 6, 3.0, style)
        add_text(ms, f"BOOST : {bv}V", col2_x + 2, y - 6, 3.0, style)

    y = y_bottom

    # OUTPUT CURRENT
    current_row_h = row_h * 1.5 if mode == 'dualsf' else row_h
    y_bottom2 = y - current_row_h
    add_rect(ms, ux1, y_bottom2, ux2, y)
    
    if mode == 'dualsf':
        add_text(ms, "OUTPUT CURRENT", ux1 + 3, y - 9, 3.0, style)
    else:
        add_text(ms, "OUTPUT CURRENT", ux1 + 3, y - 6, 3.0, style)

    if mode == 'dual':
        float_y = y - 4
        boost_y = y - 8
        ch1_fc = config.get('ch1_float_current', 20.0)
        ch1_bc = config.get('ch1_boost_current', 20.0)
        ch2_fc = config.get('ch2_float_current', 20.0)
        ch2_bc = config.get('ch2_boost_current', 20.0)
        
        add_text(ms, f"FLOAT : {ch1_fc}A", col1_x + 2, float_y, 3.0, style)
        add_text(ms, f"BOOST : {ch1_bc}A", col1_x + 2, boost_y, 3.0, style)
        add_text(ms, f"FLOAT : {ch2_fc}A", col2_x + 2, float_y, 3.0, style)
        add_text(ms, f"BOOST : {ch2_bc}A", col2_x + 2, boost_y, 3.0, style)
        
    elif mode == 'dualsf':
        float_y = y - 4.5
        ch1_fc = config.get('ch1_float_current', 100.0)
        ch1_bs = config.get('ch1_boost_start', 60.0)
        ch1_bf = config.get('ch1_boost_finish', 30.0)
        ch2_fc = config.get('ch2_float_current', 100.0)
        ch2_bs = config.get('ch2_boost_start', 60.0)
        ch2_bf = config.get('ch2_boost_finish', 30.0)
        
        # Left column
        add_text(ms, f"FLOAT : {ch1_fc}A", col1_x + 2, float_y, 3.0, style)
        line_y = float_y - 2.0
        line_start_x = col1_x - 2
        line_end_x = col1_x + 97
        v2 = make_safearray_3d([(line_start_x, line_y, 0), (line_end_x, line_y, 0)])
        ms.AddPolyline(v2)
        
        boost_label_x = col1_x - 0.5
        boost_y = y - 12
        add_text(ms, "BOOST", boost_label_x, boost_y, 3.0, style)
        
        v_line_x = boost_label_x + 15
        v_line_top = y - 6.5
        v_line_bottom = y - 15.0
        v3 = make_safearray_3d([(v_line_x, v_line_top, 0), (v_line_x, v_line_bottom, 0)])
        ms.AddPolyline(v3)
        
        start_x = boost_label_x + 35
        start_y = y - 10
        finish_y = y - 14
        sf_size = 2.4
        add_text(ms, f"START : {ch1_bs}A", start_x + shift_left, start_y, sf_size, style)
        add_text(ms, f"FINISH : {ch1_bf}A", start_x + shift_left, finish_y, sf_size, style)

        # Right column
        add_text(ms, f"FLOAT : {ch2_fc}A", col2_x + 2, float_y, 3.0, style)
        boost_label_x2 = col2_x - 2
        add_text(ms, "BOOST", boost_label_x2, boost_y, 3.0, style)
        
        v_line_x2 = boost_label_x2 + 15
        v4 = make_safearray_3d([(v_line_x2, v_line_top, 0), (v_line_x2, v_line_bottom, 0)])
        ms.AddPolyline(v4)
        
        start_x2 = boost_label_x2 + 35
        add_text(ms, f"START : {ch2_bs}A", start_x2 + shift_left, start_y, sf_size, style)
        add_text(ms, f"FINISH : {ch2_bf}A", start_x2 + shift_left, finish_y, sf_size, style)
        
    elif mode == 'ffcb':
        float_charger_c = config.get('float_charger_current', 15.0)
        fcb_fc = config.get('fcb_float_current', 15.0)
        fcb_bc = config.get('fcb_boost_current', 15.0)
        
        add_text(ms, f"{float_charger_c}A", col1_x + 2, y - 6, 3.0, style)
        float_y = y - 4
        boost_y = y - 8
        add_text(ms, f"FLOAT : {fcb_fc}A", col2_x + 2, float_y, 3.0, style)
        add_text(ms, f"BOOST : {fcb_bc}A", col2_x + 2, boost_y, 3.0, style)
        
    else:  # single
        fc = config.get('float_current', 20.0)
        bc = config.get('boost_current', 20.0)
        add_text(ms, f"FLOAT : {fc}A", col1_x + 2, y - 6, 3.0, style)
        add_text(ms, f"BOOST : {bc}A", col2_x + 2, y - 6, 3.0, style)

    y = y_bottom2
    three_bottom = y

    # Vertical lines for three-column block
    v2_x = col2_x - 4.0
    v2 = make_safearray_3d([(v2_x, three_top, 0),(v2_x, three_bottom, 0)])
    ms.AddPolyline(v2)

    # Get current year and FY range
    year = config.get('year', datetime.now().year)
    yy1 = year % 100
    yy2 = (year + 1) % 100
    fy_range = f"{yy1:02d}-{yy2:02d}"

    # SL NO
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, "SL. NO.", ux1 + 3, y - 6, 3.0, style)

    project_no = config.get('project_no', 1077)
    order_no = config.get('order_no', 2111)
    serial_no = f"LL/{fy_range}/{project_no}-OP{order_no}/BCH"
    add_text(ms, serial_no, data_x, y - 6, 3.0, style)
    y = y_bottom

    # YEAR
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, "YEAR OF MFG.", ux1 + 3, y - 6, 3.0, style)
    add_text(ms, str(year), data_x, y - 6, 3.0, style)
    y = y_bottom

    # FIRST COLUMN VERTICAL LINE
    v_line = make_safearray_3d([(vx, uy2, 0),(vx, y, 0)])
    ms.AddPolyline(v_line)

    # Footer block
    FOOT_H = 25
    y_footer_top = y
    y_footer_bottom = y_footer_top - FOOT_H

    add_line(ms, ux1, y_footer_top, ux2, y_footer_top)
    add_line(ms, ux1, y_footer_bottom, ux1, y_footer_top)
    add_line(ms, ux2, y_footer_bottom, ux2, y_footer_top)

    footer_title_h = 3.2
    footer_text_h = 2.6
    footer_text_h_a = 2.3

    fx = ux1 + 3

    # LOGO BOX
    lx1 = ux2 - logo_w + 3
    ly1 = y_footer_bottom - 2
    lx2 = lx1 + logo_w - 3
    ly2 = ly1 + logo_h

    if draw_logo_box:
        add_rect(ms, lx1, ly1, lx2, ly2)

    add_mtext(ms, r"\fConsolas|b1;LIVELINE ELECTRONICS", fx, y_footer_top - 4, 200, footer_title_h, style)
    add_text(ms, "North Ramchandrapur, Narendrapur, Kolkata : 700103", fx, y_footer_top - 12, footer_text_h_a, style)
    add_text(ms, f"{align_label('Telefax')}", fx, y_footer_top - 17, footer_text_h, style)
    add_text(ms, f"{align_label(':')}", fx+15, y_footer_top - 17, footer_text_h, style)
    add_text(ms, f"{align_label('033 2477 2094')}", fx+25, y_footer_top - 17, footer_text_h, style)
    add_text(ms, f"{align_label('Email')}", fx, y_footer_top - 22, footer_text_h, style)
    add_text(ms, f"{align_label(':')}", fx+15, y_footer_top - 22, footer_text_h, style)
    add_text(ms, f"{align_label('info@livelineindia.com')}", fx+25, y_footer_top - 22, footer_text_h, style)

    logo_block = os.path.abspath("liveline_logo.dwg")
    insert_scaled_block(ms, logo_block, lx1 - 4, ly1 + 1, logo_w, logo_h)

    # Re-draw outer and inner frames
    try:
        add_rect(ms, ox, outer_bottom, ox + w, outer_top)
        add_rect(ms, ux1, uy1, ux2, uy2)
    except Exception:
        pass

    # Zoom extents
    try:
        doc.SendCommand("_ZOOM _E ")
    except Exception:
        pass

    print("Done. Rating plate generated successfully!")

# -----------------------------
# PyQt6 GUI
# -----------------------------

class RatingPlateGUI(QMainWindow):
    # Signal emitted when background GitHub check completes: (tag, html_url, err)
    release_check_finished = pyqtSignal(object, object, object)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Battery Charger Rating Plate Generator")
        self.setMinimumSize(900, 700)
        
        # Settings
        self.auto_open_acad = False
        self.github_repo = "aamitn/winhider"  # set to your repo like 'owner/repo'
        
        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Title
        title = QLabel("Battery Charger Rating Plate Generator")
        title_font = QFont("Arial", 16, QFont.Weight.Bold)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title)
        
        # Scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        scroll.setWidget(scroll_widget)
        main_layout.addWidget(scroll)
        
        # Form sections
        scroll_layout.addWidget(self.create_general_settings())
        scroll_layout.addWidget(self.create_mode_selection())
        scroll_layout.addWidget(self.create_product_info())
        # Additional product suffix controls: checkbox + suffix text
        scroll_layout.addWidget(self.create_product_suffix_group())
        scroll_layout.addWidget(self.create_input_voltage())
        
        self.voltage_group = self.create_output_voltage()
        scroll_layout.addWidget(self.voltage_group)
        
        self.current_group = self.create_output_current()
        scroll_layout.addWidget(self.current_group)
        
        scroll_layout.addWidget(self.create_dimension_settings())
        scroll_layout.addStretch()
        
        # Generate button
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.generate_btn = QPushButton("Generate Rating Plate")
        self.generate_btn.setMinimumHeight(40)
        self.generate_btn.setStyleSheet("""
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
        self.generate_btn.clicked.connect(self.generate_plate)
        button_layout.addWidget(self.generate_btn)
        button_layout.addStretch()
        main_layout.addLayout(button_layout)
        
        self.update_voltage_current_fields()
        
        # Menu bar: Settings and Help
        menubar = self.menuBar()

        settings_menu = menubar.addMenu("Settings")
        self.auto_open_action = QAction("Auto open AutoCAD", self, checkable=True)
        self.auto_open_action.setChecked(self.auto_open_acad)
        self.auto_open_action.toggled.connect(self.set_auto_open_acad)
        settings_menu.addAction(self.auto_open_action)

        help_menu = menubar.addMenu("Help")
        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)

        self.check_action = QAction("Check GitHub Latest Release", self)
        self.check_action.triggered.connect(self.check_github_latest_release)
        help_menu.addAction(self.check_action)
    
        # Connect signal for background release check
        self.release_check_finished.connect(self._display_release_info)
    def create_general_settings(self):
        group = QGroupBox("General Settings")
        layout = QGridLayout()
        
        layout.addWidget(QLabel("Project Number (Job):"), 0, 0)
        self.project_no = QSpinBox()
        self.project_no.setRange(1, 9999)
        self.project_no.setValue(1077)
        self.project_no.setMinimumWidth(150)
        layout.addWidget(self.project_no, 0, 1)
        
        layout.addWidget(QLabel("Order Number (OP):"), 0, 2)
        self.order_no = QSpinBox()
        self.order_no.setRange(1, 9999)
        self.order_no.setValue(2111)
        self.order_no.setMinimumWidth(150)
        layout.addWidget(self.order_no, 0, 3)
        
        layout.addWidget(QLabel("Year of Manufacturing:"), 1, 0)
        self.year = QSpinBox()
        self.year.setRange(2000, 2100)
        self.year.setValue(datetime.now().year)
        self.year.setMinimumWidth(150)
        layout.addWidget(self.year, 1, 1)
        
        group.setLayout(layout)
        return group
    
    def create_mode_selection(self):
        group = QGroupBox("Charger Mode")
        layout = QHBoxLayout()
        
        layout.addWidget(QLabel("Select Mode:"))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["Single", "Dual", "FFCB", "Dual Start-Finish"])
        self.mode_combo.setMinimumWidth(200)
        self.mode_combo.currentTextChanged.connect(self.update_voltage_current_fields)
        layout.addWidget(self.mode_combo)
        layout.addStretch()
        
        group.setLayout(layout)
        return group
    
    def create_product_info(self):
        group = QGroupBox("Product Information")
        layout = QGridLayout()
        
        # Battery Voltage
        layout.addWidget(QLabel("Battery Voltage (V):"), 0, 0)
        self.battery_voltage = QDoubleSpinBox()
        self.battery_voltage.setRange(0, 999)
        self.battery_voltage.setValue(48.0)
        self.battery_voltage.setSuffix(" V")
        self.battery_voltage.valueChanged.connect(self.update_product_description)
        layout.addWidget(self.battery_voltage, 0, 1)
        
        # Battery Capacity
        layout.addWidget(QLabel("Battery Capacity (Ah):"), 0, 2)
        self.battery_capacity = QSpinBox()
        self.battery_capacity.setRange(1, 9999)
        self.battery_capacity.setValue(500)
        self.battery_capacity.setSuffix(" Ah")
        self.battery_capacity.valueChanged.connect(self.update_product_description)
        layout.addWidget(self.battery_capacity, 0, 3)
        
        # Battery Type
        layout.addWidget(QLabel("Battery Type:"), 1, 0)
        self.battery_type = QComboBox()
        self.battery_type.addItems([
            "LEAD ACID BATTERY",
            "VRLA BATTERY",
            "TUBULAR BATTERY",
            "SMF BATTERY",
            "GEL BATTERY",
            "AGM BATTERY",
            "LITHIUM ION BATTERY",
            "Ni-Cd BATTERY"
        ])
        self.battery_type.currentTextChanged.connect(self.update_product_description)
        layout.addWidget(self.battery_type, 1, 1, 1, 3)
        
        # Generated Product Description (Read-only display)
        layout.addWidget(QLabel("Generated Description:"), 2, 0, 1, 4)
        self.product_desc = QLineEdit()
        self.product_desc.setReadOnly(True)
        self.product_desc.setStyleSheet("""
            QLineEdit {
                background-color: rgba(76, 175, 80, 0.05);
                border: 1px solid #ccc;
                padding: 5px;
                font-weight: bold;
            }
        """)
        layout.addWidget(self.product_desc, 3, 0, 1, 4)

        # Product font height control (allows adjusting product MText size)
        layout.addWidget(QLabel("Product Font Height:"), 4, 0)
        self.product_font_h = QDoubleSpinBox()
        self.product_font_h.setRange(0.5, 20.0)
        self.product_font_h.setSingleStep(0.1)
        self.product_font_h.setValue(3.8)
        self.product_font_h.setSuffix(" mm")
        self.product_font_h.valueChanged.connect(self.update_product_description)
        layout.addWidget(self.product_font_h, 4, 1)
        
        group.setLayout(layout)
        
        # Initialize product description
        self.update_product_description()
        
        return group
    
    def get_charger_current(self):
        """Get the maximum charger current based on mode and input values"""
        mode = self.mode_combo.currentText()
        
        try:
            if mode == "Single":
                # Get max of float or boost current
                float_curr = getattr(self, 'float_current', None)
                boost_curr = getattr(self, 'boost_current', None)
                if float_curr and boost_curr:
                    return max(float_curr.value(), boost_curr.value())
                return 100.0  # Default
                
            elif mode == "Dual":
                # Sum of both chargers (max of float/boost for each)
                ch1_float = getattr(self, 'ch1_float_current', None)
                ch1_boost = getattr(self, 'ch1_boost_current', None)
                ch2_float = getattr(self, 'ch2_float_current', None)
                ch2_boost = getattr(self, 'ch2_boost_current', None)
                
                if all([ch1_float, ch1_boost, ch2_float, ch2_boost]):
                    ch1_max = max(ch1_float.value(), ch1_boost.value())
                    ch2_max = max(ch2_float.value(), ch2_boost.value())
                    return ch1_max + ch2_max
                return 100.0
                
            elif mode == "FFCB":
                # Sum of float charger and FCB charger (max of float/boost)
                float_curr = getattr(self, 'float_charger_current', None)
                fcb_float = getattr(self, 'fcb_float_current', None)
                fcb_boost = getattr(self, 'fcb_boost_current', None)
                
                if all([float_curr, fcb_float, fcb_boost]):
                    fcb_max = max(fcb_float.value(), fcb_boost.value())
                    return float_curr.value() + fcb_max
                return 100.0
                
            elif mode == "Dual Start-Finish":
                # Sum of both chargers (max of float or boost start for each)
                ch1_float = getattr(self, 'ch1_float_current_sf', None)
                ch1_start = getattr(self, 'ch1_boost_start', None)
                ch2_float = getattr(self, 'ch2_float_current_sf', None)
                ch2_start = getattr(self, 'ch2_boost_start', None)
                
                if all([ch1_float, ch1_start, ch2_float, ch2_start]):
                    ch1_max = max(ch1_float.value(), ch1_start.value())
                    ch2_max = max(ch2_float.value(), ch2_start.value())
                    return ch1_max + ch2_max
                return 100.0
        except:
            return 100.0  # Default fallback
        
        return 100.0
    
    def update_product_description(self):
        """Dynamically generate product description based on inputs"""
        mode = self.mode_combo.currentText()
        battery_v = self.battery_voltage.value()
        battery_cap = self.battery_capacity.value()
        battery_type = self.battery_type.currentText()
        
        # Get charger current
        charger_current = self.get_charger_current()
        
        # Determine charger type description
        mode_descriptions = {
            "Single": "FLOAT CUM BOOST",
            "Dual": "DUAL FLOAT CUM BOOST",
            "FFCB": "FLOAT & FLOAT CUM BOOST",
            "Dual Start-Finish": "DUAL FLOAT CUM BOOST"
        }
        
        charger_type = mode_descriptions.get(mode, "FLOAT CUM BOOST")
        
        # Build description
        description = f"{int(battery_v)}V {int(charger_current)}A {charger_type} BATTERY CHARGER FOR {battery_cap}AH {battery_type}"
        
        self.product_desc.setText(description)
    
    def create_input_voltage(self):
        group = QGroupBox("Input Voltage")
        layout = QGridLayout()

        # Voltage selection (dropdown)
        layout.addWidget(QLabel("Supply Voltage:"), 0, 0)
        self.voltage_combo = QComboBox()
        self.voltage_combo.addItems(["415", "230"])
        self.voltage_combo.setCurrentText("415")
        layout.addWidget(self.voltage_combo, 0, 1)

        # Phase display (derived from voltage)
        layout.addWidget(QLabel("Phase:"), 0, 2)
        self.phase_label = QLabel("3")
        layout.addWidget(self.phase_label, 0, 3)

        # Wires count (editable for multi-phase, fixed for 230)
        layout.addWidget(QLabel("Number of Wires:"), 1, 0)
        self.wires_spin = QSpinBox()
        self.wires_spin.setRange(1, 6)
        self.wires_spin.setValue(4)
        layout.addWidget(self.wires_spin, 1, 1)

        # Frequency (Hz)
        layout.addWidget(QLabel("Frequency (Hz):"), 1, 2)
        self.freq_spin = QSpinBox()
        self.freq_spin.setRange(40, 60)
        self.freq_spin.setValue(50)
        layout.addWidget(self.freq_spin, 1, 3)

        # Constructed read-only input-voltage string shown to user
        layout.addWidget(QLabel("Constructed Input String:"), 2, 0)
        self.input_voltage = QLineEdit()
        self.input_voltage.setReadOnly(True)
        layout.addWidget(self.input_voltage, 2, 1, 1, 3)

        # Wire up signals
        self.voltage_combo.currentTextChanged.connect(self.on_voltage_changed)
        self.wires_spin.valueChanged.connect(self.update_input_voltage_display)
        self.freq_spin.valueChanged.connect(self.update_input_voltage_display)

        # Initialize display
        self.on_voltage_changed(self.voltage_combo.currentText())

        group.setLayout(layout)
        return group

    def create_product_suffix_group(self):
        group = QGroupBox("Product Suffix")
        layout = QHBoxLayout()

        self.append_suffix_checkbox = QCheckBox("DBDB Integrated")
        self.append_suffix_checkbox.setChecked(False)
        layout.addWidget(self.append_suffix_checkbox)

        self.suffix_edit = QLineEdit()
        self.suffix_edit.setText(" WITH INTEGRATED DCDB")
        self.suffix_edit.setPlaceholderText("Suffix to append when checkbox is checked")
        self.suffix_edit.setVisible(False)
        layout.addWidget(self.suffix_edit)

        # Toggle visibility of suffix edit
        self.append_suffix_checkbox.toggled.connect(lambda v: self.suffix_edit.setVisible(v))

        group.setLayout(layout)
        return group

    def on_voltage_changed(self, val):
        """Handle changes to voltage dropdown: set phase and wires behavior."""
        try:
            v = int(val)
        except Exception:
            v = 415

        if v == 230:
            phase = 2
            # For 230, wires fixed to 2
            self.wires_spin.setValue(2)
            self.wires_spin.setEnabled(False)
        else:
            phase = 3
            # For 415 (3-phase) allow wires selection (default 4)
            if not self.wires_spin.isEnabled():
                self.wires_spin.setEnabled(True)
                self.wires_spin.setValue(4)

        self.phase_label.setText(str(phase))
        self.update_input_voltage_display()

    def set_auto_open_acad(self, checked: bool):
        self.auto_open_acad = bool(checked)

    def show_about_dialog(self):
        text = (
            "Plategen - Battery Charger Rating Plate Generator\n"
            "Author: Liveline (placeholder)\n"
            f"GitHub repo: {self.github_repo}\n"
        )
        QMessageBox.information(self, "About", text)

    def check_github_latest_release(self):
        repo = self.github_repo
        if not repo or '/' not in repo:
            QMessageBox.warning(self, "Check Latest Release", "No GitHub repo configured. Set `self.github_repo`.")
            return

        # Disable the action and show status while checking
        try:
            self.check_action.setEnabled(False)
        except Exception:
            pass
        try:
            self.statusBar().showMessage("Checking GitHub latest release...")
        except Exception:
            pass

        # Run network fetch in a background thread to avoid blocking UI
        def worker():
            try:
                tag, html_url, err = fetch_latest_github_release(repo)
            except Exception as e:
                tag, html_url, err = None, None, str(e)

            # Emit signal to schedule UI update on main thread
            try:
                self.release_check_finished.emit(tag, html_url, err)
            except Exception:
                # Fallback: directly call UI handler on main thread via QTimer
                try:
                    from PyQt6.QtCore import QTimer
                    QTimer.singleShot(0, lambda: self._display_release_info(tag, html_url, err))
                except Exception:
                    # As last resort, call directly (may block)
                    self._display_release_info(tag, html_url, err)

        t = threading.Thread(target=worker, daemon=True)
        t.start()

    def _display_release_info(self, tag, html_url, err):
        """Display release info in the main thread and offer to open the release URL if newer."""
        # Re-enable the check action and clear status
        try:
            self.check_action.setEnabled(True)
        except Exception:
            pass
        try:
            self.statusBar().clearMessage()
        except Exception:
            pass

        if err:
            QMessageBox.critical(self, "Check Latest Release", f"Failed to fetch latest release: {err}")
            return

        if not tag:
            QMessageBox.information(self, "Latest Release", "No release information found.")
            return

        cmp = compare_versions(APP_VERSION, tag)
        if cmp < 0:
            # remote is newer
            rv = QMessageBox(self)
            rv.setIcon(QMessageBox.Icon.Information)
            rv.setWindowTitle("Update Available")
            rv.setText(f"A newer release is available: {tag}\nYou have: {APP_VERSION}")
            open_btn = rv.addButton("Download Update", QMessageBox.ButtonRole.AcceptRole)
            rv.addButton("Dismiss", QMessageBox.ButtonRole.RejectRole)
            rv.exec()
            if rv.clickedButton() == open_btn and html_url:
                webbrowser.open(html_url)
        else:
            QMessageBox.information(self, "Latest Release", f"You are up-to-date. Latest: {tag} (You: {APP_VERSION})")

    def update_input_voltage_display(self):
        """Construct the human-readable input-voltage string from controls."""
        try:
            voltage = int(self.voltage_combo.currentText())
        except Exception:
            voltage = 415

        phase = int(self.phase_label.text()) if hasattr(self, 'phase_label') else (3 if voltage == 415 else 2)
        wires = int(self.wires_spin.value()) if hasattr(self, 'wires_spin') else (4 if phase == 3 else 2)
        freq = int(self.freq_spin.value()) if hasattr(self, 'freq_spin') else 50

        s = f"{voltage}V AC, {phase} PHASE, {wires} WIRES, {freq}HZ"
        self.input_voltage.setText(s)
    
    def create_output_voltage(self):
        group = QGroupBox("Output Voltage")
        self.voltage_layout = QGridLayout()
        group.setLayout(self.voltage_layout)
        return group
    
    def create_output_current(self):
        group = QGroupBox("Output Current")
        self.current_layout = QGridLayout()
        group.setLayout(self.current_layout)
        return group
    
    def create_dimension_settings(self):
        group = QGroupBox("Plate Dimensions (mm)")
        layout = QGridLayout()
        
        layout.addWidget(QLabel("Plate Width:"), 0, 0)
        self.plate_width = QDoubleSpinBox()
        self.plate_width.setRange(50, 500)
        self.plate_width.setValue(150.0)
        self.plate_width.setSuffix(" mm")
        layout.addWidget(self.plate_width, 0, 1)
        
        layout.addWidget(QLabel("Plate Height:"), 0, 2)
        self.plate_height = QDoubleSpinBox()
        self.plate_height.setRange(50, 500)
        self.plate_height.setValue(100.0)
        self.plate_height.setSuffix(" mm")
        layout.addWidget(self.plate_height, 0, 3)
        
        layout.addWidget(QLabel("Margin:"), 1, 0)
        self.margin = QDoubleSpinBox()
        self.margin.setRange(0, 20)
        self.margin.setValue(3.0)
        self.margin.setSuffix(" mm")
        layout.addWidget(self.margin, 1, 1)
        
        layout.addWidget(QLabel("Logo Width:"), 1, 2)
        self.logo_width = QDoubleSpinBox()
        self.logo_width.setRange(10, 100)
        self.logo_width.setValue(40.0)
        self.logo_width.setSuffix(" mm")
        layout.addWidget(self.logo_width, 1, 3)
        
        layout.addWidget(QLabel("Logo Height:"), 2, 0)
        self.logo_height = QDoubleSpinBox()
        self.logo_height.setRange(10, 100)
        self.logo_height.setValue(30.0)
        self.logo_height.setSuffix(" mm")
        layout.addWidget(self.logo_height, 2, 1)
        
        layout.addWidget(QLabel("Dimension Text Size:"), 2, 2)
        self.dim_text_size = QDoubleSpinBox()
        self.dim_text_size.setRange(1, 20)
        self.dim_text_size.setValue(5.0)
        layout.addWidget(self.dim_text_size, 2, 3)
        
        # Use Scale Checkbox
        self.use_scale_checkbox = QCheckBox("Use 1:1 Scale (Show Actual Dimensions)")
        self.use_scale_checkbox.setChecked(True)
        self.use_scale_checkbox.setStyleSheet("font-weight: bold; margin-top: 10px;")
        self.use_scale_checkbox.stateChanged.connect(self.toggle_dimension_overrides)
        layout.addWidget(self.use_scale_checkbox, 3, 0, 1, 4)
        
        # Dimension Override Section
        override_label = QLabel("Dimension Text Overrides")
        override_label.setStyleSheet("font-weight: bold; margin-top: 5px; color: #666;")
        layout.addWidget(override_label, 4, 0, 1, 4)
        
        layout.addWidget(QLabel("Override Width Text:"), 5, 0)
        self.dim_width_override = QDoubleSpinBox()
        self.dim_width_override.setRange(0, 999)
        self.dim_width_override.setValue(150.0)
        self.dim_width_override.setSuffix(" mm")
        self.dim_width_override.setEnabled(False)  # Disabled by default
        layout.addWidget(self.dim_width_override, 5, 1)
        
        layout.addWidget(QLabel("Override Height Text:"), 5, 2)
        self.dim_height_override = QDoubleSpinBox()
        self.dim_height_override.setRange(0, 999)
        self.dim_height_override.setValue(100.0)
        self.dim_height_override.setSuffix(" mm")
        self.dim_height_override.setEnabled(False)  # Disabled by default
        layout.addWidget(self.dim_height_override, 5, 3)
        
        group.setLayout(layout)
        return group
    
    def toggle_dimension_overrides(self, state):
        """Enable/disable dimension overrides based on scale checkbox"""
        use_scale = self.use_scale_checkbox.isChecked()
        
        # When scale is checked (1:1), disable overrides
        # When scale is unchecked, enable overrides
        self.dim_width_override.setEnabled(not use_scale)
        self.dim_height_override.setEnabled(not use_scale)
        
        # Visual feedback with dark mode compatible styling
        if use_scale:
            self.dim_width_override.setStyleSheet("")
            self.dim_height_override.setStyleSheet("")
        else:
            # Use border and font styling that works in both light and dark modes
            override_style = """
                QDoubleSpinBox {
                    border: 2px solid #4CAF50;
                    font-weight: bold;
                }
                QDoubleSpinBox:enabled {
                    background-color: rgba(76, 175, 80, 0.1);
                }
            """
            self.dim_width_override.setStyleSheet(override_style)
            self.dim_height_override.setStyleSheet(override_style)
    
    def clear_layout(self, layout):
        while layout.count():
            child = layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
    
    def update_voltage_current_fields(self):
        mode = self.mode_combo.currentText()
        self.clear_layout(self.voltage_layout)
        self.clear_layout(self.current_layout)
        
        if mode == "Single":
            self.create_single_mode_fields()
        elif mode == "Dual":
            self.create_dual_mode_fields()
        elif mode == "FFCB":
            self.create_ffcb_mode_fields()
        elif mode == "Dual Start-Finish":
            self.create_dual_sf_mode_fields()
        
        # Update product description when mode changes
        self.update_product_description()
    
    def create_single_mode_fields(self):
        self.voltage_layout.addWidget(QLabel("Float Voltage (V):"), 0, 0)
        self.float_voltage = QDoubleSpinBox()
        self.float_voltage.setRange(0, 999)
        self.float_voltage.setValue(123.75)
        self.voltage_layout.addWidget(self.float_voltage, 0, 1)
        
        self.voltage_layout.addWidget(QLabel("Boost Voltage (V):"), 0, 2)
        self.boost_voltage = QDoubleSpinBox()
        self.boost_voltage.setRange(0, 999)
        self.boost_voltage.setValue(126.5)
        self.voltage_layout.addWidget(self.boost_voltage, 0, 3)
        
        self.current_layout.addWidget(QLabel("Float Current (A):"), 0, 0)
        self.float_current = QDoubleSpinBox()
        self.float_current.setRange(0, 999)
        self.float_current.setValue(20.0)
        self.float_current.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.float_current, 0, 1)
        
        self.current_layout.addWidget(QLabel("Boost Current (A):"), 0, 2)
        self.boost_current = QDoubleSpinBox()
        self.boost_current.setRange(0, 999)
        self.boost_current.setValue(20.0)
        self.boost_current.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.boost_current, 0, 3)
    
    def create_dual_mode_fields(self):
        self.voltage_layout.addWidget(QLabel("CHARGER-I", font=QFont("Arial", 10, QFont.Weight.Bold)), 0, 0, 1, 2)
        self.voltage_layout.addWidget(QLabel("Float Voltage (V):"), 1, 0)
        self.ch1_float_voltage = QDoubleSpinBox()
        self.ch1_float_voltage.setRange(0, 999)
        self.ch1_float_voltage.setValue(123.75)
        self.voltage_layout.addWidget(self.ch1_float_voltage, 1, 1)
        
        self.voltage_layout.addWidget(QLabel("Boost Voltage (V):"), 2, 0)
        self.ch1_boost_voltage = QDoubleSpinBox()
        self.ch1_boost_voltage.setRange(0, 999)
        self.ch1_boost_voltage.setValue(126.5)
        self.voltage_layout.addWidget(self.ch1_boost_voltage, 2, 1)
        
        self.voltage_layout.addWidget(QLabel("CHARGER-II", font=QFont("Arial", 10, QFont.Weight.Bold)), 0, 2, 1, 2)
        self.voltage_layout.addWidget(QLabel("Float Voltage (V):"), 1, 2)
        self.ch2_float_voltage = QDoubleSpinBox()
        self.ch2_float_voltage.setRange(0, 999)
        self.ch2_float_voltage.setValue(123.75)
        self.voltage_layout.addWidget(self.ch2_float_voltage, 1, 3)
        
        self.voltage_layout.addWidget(QLabel("Boost Voltage (V):"), 2, 2)
        self.ch2_boost_voltage = QDoubleSpinBox()
        self.ch2_boost_voltage.setRange(0, 999)
        self.ch2_boost_voltage.setValue(126.5)
        self.voltage_layout.addWidget(self.ch2_boost_voltage, 2, 3)
        
        self.current_layout.addWidget(QLabel("CHARGER-I", font=QFont("Arial", 10, QFont.Weight.Bold)), 0, 0, 1, 2)
        self.current_layout.addWidget(QLabel("Float Current (A):"), 1, 0)
        self.ch1_float_current = QDoubleSpinBox()
        self.ch1_float_current.setRange(0, 999)
        self.ch1_float_current.setValue(20.0)
        self.ch1_float_current.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.ch1_float_current, 1, 1)
        
        self.current_layout.addWidget(QLabel("Boost Current (A):"), 2, 0)
        self.ch1_boost_current = QDoubleSpinBox()
        self.ch1_boost_current.setRange(0, 999)
        self.ch1_boost_current.setValue(20.0)
        self.ch1_boost_current.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.ch1_boost_current, 2, 1)
        
        self.current_layout.addWidget(QLabel("CHARGER-II", font=QFont("Arial", 10, QFont.Weight.Bold)), 0, 2, 1, 2)
        self.current_layout.addWidget(QLabel("Float Current (A):"), 1, 2)
        self.ch2_float_current = QDoubleSpinBox()
        self.ch2_float_current.setRange(0, 999)
        self.ch2_float_current.setValue(20.0)
        self.ch2_float_current.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.ch2_float_current, 1, 3)
        
        self.current_layout.addWidget(QLabel("Boost Current (A):"), 2, 2)
        self.ch2_boost_current = QDoubleSpinBox()
        self.ch2_boost_current.setRange(0, 999)
        self.ch2_boost_current.setValue(20.0)
        self.ch2_boost_current.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.ch2_boost_current, 2, 3)
    
    def create_ffcb_mode_fields(self):
        self.voltage_layout.addWidget(QLabel("FLOAT CHARGER", font=QFont("Arial", 10, QFont.Weight.Bold)), 0, 0, 1, 2)
        self.voltage_layout.addWidget(QLabel("Voltage (V):"), 1, 0)
        self.float_charger_voltage = QDoubleSpinBox()
        self.float_charger_voltage.setRange(0, 999)
        self.float_charger_voltage.setValue(123.75)
        self.voltage_layout.addWidget(self.float_charger_voltage, 1, 1)
        
        self.voltage_layout.addWidget(QLabel("FCB CHARGER", font=QFont("Arial", 10, QFont.Weight.Bold)), 0, 2, 1, 2)
        self.voltage_layout.addWidget(QLabel("Float Voltage (V):"), 1, 2)
        self.fcb_float_voltage = QDoubleSpinBox()
        self.fcb_float_voltage.setRange(0, 999)
        self.fcb_float_voltage.setValue(123.75)
        self.voltage_layout.addWidget(self.fcb_float_voltage, 1, 3)
        
        self.voltage_layout.addWidget(QLabel("Boost Voltage (V):"), 2, 2)
        self.fcb_boost_voltage = QDoubleSpinBox()
        self.fcb_boost_voltage.setRange(0, 999)
        self.fcb_boost_voltage.setValue(126.5)
        self.voltage_layout.addWidget(self.fcb_boost_voltage, 2, 3)
        
        self.current_layout.addWidget(QLabel("FLOAT CHARGER", font=QFont("Arial", 10, QFont.Weight.Bold)), 0, 0, 1, 2)
        self.current_layout.addWidget(QLabel("Current (A):"), 1, 0)
        self.float_charger_current = QDoubleSpinBox()
        self.float_charger_current.setRange(0, 999)
        self.float_charger_current.setValue(15.0)
        self.float_charger_current.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.float_charger_current, 1, 1)
        
        self.current_layout.addWidget(QLabel("FCB CHARGER", font=QFont("Arial", 10, QFont.Weight.Bold)), 0, 2, 1, 2)
        self.current_layout.addWidget(QLabel("Float Current (A):"), 1, 2)
        self.fcb_float_current = QDoubleSpinBox()
        self.fcb_float_current.setRange(0, 999)
        self.fcb_float_current.setValue(15.0)
        self.fcb_float_current.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.fcb_float_current, 1, 3)
        
        self.current_layout.addWidget(QLabel("Boost Current (A):"), 2, 2)
        self.fcb_boost_current = QDoubleSpinBox()
        self.fcb_boost_current.setRange(0, 999)
        self.fcb_boost_current.setValue(15.0)
        self.fcb_boost_current.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.fcb_boost_current, 2, 3)
    
    def create_dual_sf_mode_fields(self):
        self.voltage_layout.addWidget(QLabel("CHARGER-I", font=QFont("Arial", 10, QFont.Weight.Bold)), 0, 0, 1, 2)
        self.voltage_layout.addWidget(QLabel("Float Voltage (V):"), 1, 0)
        self.ch1_float_voltage_sf = QDoubleSpinBox()
        self.ch1_float_voltage_sf.setRange(0, 999)
        self.ch1_float_voltage_sf.setValue(54.0)
        self.voltage_layout.addWidget(self.ch1_float_voltage_sf, 1, 1)
        
        self.voltage_layout.addWidget(QLabel("Boost Voltage (V):"), 2, 0)
        self.ch1_boost_voltage_sf = QDoubleSpinBox()
        self.ch1_boost_voltage_sf.setRange(0, 999)
        self.ch1_boost_voltage_sf.setValue(66.0)
        self.voltage_layout.addWidget(self.ch1_boost_voltage_sf, 2, 1)
        
        self.voltage_layout.addWidget(QLabel("CHARGER-II", font=QFont("Arial", 10, QFont.Weight.Bold)), 0, 2, 1, 2)
        self.voltage_layout.addWidget(QLabel("Float Voltage (V):"), 1, 2)
        self.ch2_float_voltage_sf = QDoubleSpinBox()
        self.ch2_float_voltage_sf.setRange(0, 999)
        self.ch2_float_voltage_sf.setValue(54.0)
        self.voltage_layout.addWidget(self.ch2_float_voltage_sf, 1, 3)
        
        self.voltage_layout.addWidget(QLabel("Boost Voltage (V):"), 2, 2)
        self.ch2_boost_voltage_sf = QDoubleSpinBox()
        self.ch2_boost_voltage_sf.setRange(0, 999)
        self.ch2_boost_voltage_sf.setValue(66.0)
        self.voltage_layout.addWidget(self.ch2_boost_voltage_sf, 2, 3)
        
        self.current_layout.addWidget(QLabel("CHARGER-I", font=QFont("Arial", 10, QFont.Weight.Bold)), 0, 0, 1, 2)
        self.current_layout.addWidget(QLabel("Float Current (A):"), 1, 0)
        self.ch1_float_current_sf = QDoubleSpinBox()
        self.ch1_float_current_sf.setRange(0, 999)
        self.ch1_float_current_sf.setValue(100.0)
        self.ch1_float_current_sf.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.ch1_float_current_sf, 1, 1)
        
        self.current_layout.addWidget(QLabel("Boost Start (A):"), 2, 0)
        self.ch1_boost_start = QDoubleSpinBox()
        self.ch1_boost_start.setRange(0, 999)
        self.ch1_boost_start.setValue(60.0)
        self.ch1_boost_start.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.ch1_boost_start, 2, 1)
        
        self.current_layout.addWidget(QLabel("Boost Finish (A):"), 3, 0)
        self.ch1_boost_finish = QDoubleSpinBox()
        self.ch1_boost_finish.setRange(0, 999)
        self.ch1_boost_finish.setValue(30.0)
        self.ch1_boost_finish.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.ch1_boost_finish, 3, 1)
        
        self.current_layout.addWidget(QLabel("CHARGER-II", font=QFont("Arial", 10, QFont.Weight.Bold)), 0, 2, 1, 2)
        self.current_layout.addWidget(QLabel("Float Current (A):"), 1, 2)
        self.ch2_float_current_sf = QDoubleSpinBox()
        self.ch2_float_current_sf.setRange(0, 999)
        self.ch2_float_current_sf.setValue(100.0)
        self.ch2_float_current_sf.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.ch2_float_current_sf, 1, 3)
        
        self.current_layout.addWidget(QLabel("Boost Start (A):"), 2, 2)
        self.ch2_boost_start = QDoubleSpinBox()
        self.ch2_boost_start.setRange(0, 999)
        self.ch2_boost_start.setValue(60.0)
        self.ch2_boost_start.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.ch2_boost_start, 2, 3)
        
        self.current_layout.addWidget(QLabel("Boost Finish (A):"), 3, 2)
        self.ch2_boost_finish = QDoubleSpinBox()
        self.ch2_boost_finish.setRange(0, 999)
        self.ch2_boost_finish.setValue(30.0)
        self.ch2_boost_finish.valueChanged.connect(self.update_product_description)
        self.current_layout.addWidget(self.ch2_boost_finish, 3, 3)
    
    def get_config(self):
        mode = self.mode_combo.currentText()
        mode_map = {
            "Single": "single",
            "Dual": "dual",
            "FFCB": "ffcb",
            "Dual Start-Finish": "dualsf"
        }
        
        config = {
            'mode': mode_map[mode],
            'project_no': self.project_no.value(),
            'order_no': self.order_no.value(),
            'year': self.year.value(),
            
            # Product description (optionally append static suffix)
            
            'product_desc': (self.product_desc.text() + self.suffix_edit.text()) if getattr(self, 'append_suffix_checkbox', None) and self.append_suffix_checkbox.isChecked() else self.product_desc.text(),
            'product_font_h': self.product_font_h.value(),
            'input_voltage': self.input_voltage.text(),
            'plate_width': self.plate_width.value(),
            'plate_height': self.plate_height.value(),
            'margin': self.margin.value(),
            'logo_width': self.logo_width.value(),
            'logo_height': self.logo_height.value(),
            'dim_text_size': self.dim_text_size.value(),
            'dim_width_override': None if self.use_scale_checkbox.isChecked() else self.dim_width_override.value(),
            'dim_height_override': None if self.use_scale_checkbox.isChecked() else self.dim_height_override.value(),
            'offset_x': 100.0,
            'offset_y': 100.0,
            'label_w': 40.0,
            'product_h': 20.0,
            'row_h': 10.0,
            'draw_logo_box': False,
        }
        
        if mode == "Single":
            config['float_voltage'] = self.float_voltage.value()
            config['boost_voltage'] = self.boost_voltage.value()
            config['float_current'] = self.float_current.value()
            config['boost_current'] = self.boost_current.value()
        elif mode == "Dual":
            config['ch1_float_voltage'] = self.ch1_float_voltage.value()
            config['ch1_boost_voltage'] = self.ch1_boost_voltage.value()
            config['ch2_float_voltage'] = self.ch2_float_voltage.value()
            config['ch2_boost_voltage'] = self.ch2_boost_voltage.value()
            config['ch1_float_current'] = self.ch1_float_current.value()
            config['ch1_boost_current'] = self.ch1_boost_current.value()
            config['ch2_float_current'] = self.ch2_float_current.value()
            config['ch2_boost_current'] = self.ch2_boost_current.value()
        elif mode == "FFCB":
            config['float_charger_voltage'] = self.float_charger_voltage.value()
            config['fcb_float_voltage'] = self.fcb_float_voltage.value()
            config['fcb_boost_voltage'] = self.fcb_boost_voltage.value()
            config['float_charger_current'] = self.float_charger_current.value()
            config['fcb_float_current'] = self.fcb_float_current.value()
            config['fcb_boost_current'] = self.fcb_boost_current.value()
        elif mode == "Dual Start-Finish":
            config['ch1_float_voltage'] = self.ch1_float_voltage_sf.value()
            config['ch1_boost_voltage'] = self.ch1_boost_voltage_sf.value()
            config['ch2_float_voltage'] = self.ch2_float_voltage_sf.value()
            config['ch2_boost_voltage'] = self.ch2_boost_voltage_sf.value()
            config['ch1_float_current'] = self.ch1_float_current_sf.value()
            config['ch1_boost_start'] = self.ch1_boost_start.value()
            config['ch1_boost_finish'] = self.ch1_boost_finish.value()
            config['ch2_float_current'] = self.ch2_float_current_sf.value()
            config['ch2_boost_start'] = self.ch2_boost_start.value()
            config['ch2_boost_finish'] = self.ch2_boost_finish.value()
        
        return config
    
    def generate_plate(self):
        try:
            config = self.get_config()
            
            pythoncom.CoInitialize()
            acad = None
            # First try to get an active AutoCAD COM object
            try:
                acad = win32com.client.GetActiveObject("AutoCAD.Application")
            except Exception:
                acad = None

            # If not available, attempt to Dispatch common AutoCAD ProgIDs
            if acad is None:
                progids = [
                    "AutoCAD.Application",
                    "AutoCAD.Application.24",
                    "AutoCAD.Application.23",
                    "AutoCAD.Application.22",
                    "AutoCAD.Application.21",
                    "AutoCADElectrical.Application",
                    "AutoCADLT.Application",
                ]
                for pid in progids:
                    try:
                        acad = win32com.client.Dispatch(pid)
                        if acad:
                            print(f"Dispatched AutoCAD via ProgID: {pid}")
                            break
                    except Exception:
                        acad = None

            # If still not found and user wants AutoCAD auto-open, try launching executables
            if acad is None and self.auto_open_acad:
                exe_names = ["acad.exe", "accoreconsole.exe", "acadlt.exe"]
                for exe in exe_names:
                    try:
                        print(f"Attempting to start executable: {exe}")
                        subprocess.Popen([exe], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                        # give it a moment to register COM server
                        time.sleep(2.0)
                        try:
                            acad = win32com.client.GetActiveObject("AutoCAD.Application")
                        except Exception:
                            acad = None
                        if acad:
                            print(f"Started AutoCAD via executable: {exe}")
                            break
                    except Exception:
                        acad = None

            if acad is None:
                # As a last resort try a generic Dispatch
                try:
                    acad = win32com.client.Dispatch("AutoCAD.Application")
                except Exception as e:
                    raise RuntimeError(f"Unable to start or connect to AutoCAD: {e}")

            # If user asked to auto-open AutoCAD, ensure there's an open drawing. Handle COM 'call was rejected' by retrying.
            doc = None
            attempts = 0
            while attempts < 8:
                attempts += 1
                try:
                    doc = acad.ActiveDocument
                    break
                except Exception as e:
                    # COM call rejected may return HRESULT -2147418111; wait and retry
                    args = getattr(e, 'args', None)
                    code = args[0] if args and len(args) > 0 else None
                    if code == -2147418111:
                        time.sleep(0.6)
                        continue
                    # If no active document and auto_open_acad requested, try to add a new document
                    if self.auto_open_acad:
                        try:
                            doc = acad.Documents.Add()
                            break
                        except Exception as e2:
                            args2 = getattr(e2, 'args', None)
                            code2 = args2[0] if args2 and len(args2) > 0 else None
                            if code2 == -2147418111:
                                time.sleep(0.6)
                                continue
                            # else raise after attempts
                    # otherwise break and raise
                    doc = None
                    break

            if doc is None:
                raise RuntimeError("No active AutoCAD document and could not create one.")

            draw_rating_plate(doc, config)
            
            QMessageBox.information(self, "Success", "Rating plate generated successfully!")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred:\n{str(e)}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

# -----------------------------
# Main
# -----------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    # app.addLibraryPath(os.path.join(os.path.dirname(sys.argv[0]), "plugins"))

    # Ensure app icon exists and set it for the application and main window
    try:
        icon_path = ensure_app_icon()
        if icon_path:
            app.setWindowIcon(QIcon(icon_path))
    except Exception:
        pass

    window = RatingPlateGUI()
    try:
        if icon_path:
            window.setWindowIcon(QIcon(icon_path))
    except Exception:
        pass

    window.show()
    sys.exit(app.exec())
import win32com.client
import pythoncom
from array import array
import os
import time

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
        # Create empty style – font must already be assigned in DWG manually
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
    # AddMText(position, width, text) — position is lower-left of the box.
    # We want top-left alignment, so we provide a point slightly below top.
    p = make_point_variant(x, y, 0)
    mt = ms.AddMText(p, float(width), text)
    mt.Height = float(height)
    mt.StyleName = style.Name
    # set attachment to top-left (2). Some AutoCAD versions accept Attachment property.
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
    text_height -> Optional float to override the font size using MText codes.
    """
    p1 = make_point_variant(x1, y1)
    p2 = make_point_variant(x2, y2)
    # The user provided a fixed offset for p_dim_alligned
    # p_dim = make_point_variant(dim_x, dim_y) # Original
    p_dim_alligned = make_point_variant(dim_x - 3 , dim_y - 3 )  # Adjusted to better center text

    # dim = ms.AddDimAligned(p1, p2, p_dim) # Original
    dim = ms.AddDimAligned(p1, p2, p_dim_alligned)

    # --- Text Placement Fixes ---
    # 1. Disable text background mask for the line to show through.
    try:
        dim.TextFill = False
    except Exception:
        pass
        
    # 2. Set Text Gap to control text position relative to the dimension line.
    # Note: The user set this to 1.5, which will place the text *above* the line.
    # If the goal is "on the line", a small negative number like -0.5 is needed. 
    # Sticking to user's 1.5 for now, but commenting the "on the line" goal for clarity.
    # 2. Set Text Gap to control text position relative to the dimension line.
    dim.TextGap = 1.5 
    # --- END Fixes ---

    # --- Fix: Apply Consolas Font and Height ---
    
    # 1. Start with the core text content: use override_text or the default measurement marker
    content = override_text if override_text is not None else "<>"
    
    # 2. Start building the MText override string, starting with Consolas font code
    final_override = r"\FConsolas;"

    if text_height:
        # 3. Apply height if specified
        final_override += r"\H{0};".format(text_height)
        
    # 4. Add the content (either override text or '<>')
    final_override += content
    
    # Apply the compiled override string
    dim.TextOverride = final_override
        
    return dim


def insert_scaled_block(ms, block_path, x, y, target_w, target_h):
    ins_pt = make_point_variant(x, y, 0)

    # Insert at scale 1.0 first
    blk = ms.InsertBlock(ins_pt, block_path, 1.0, 1.0, 1.0, 0)

    # Update extents
    try:
        blk.Update()
    except:
        pass

    # Get extents
    bb = blk.GetBoundingBox()
    (xmin, ymin, zmin) = bb[0]
    (xmax, ymax, zmax) = bb[1]

    bw = xmax - xmin
    bh = ymax - ymin

    if bw == 0 or bh == 0:
        print("Block has zero geometry.")
        return blk

    # Keep aspect ratio
    sx = target_w / bw
    sy = target_h / bh
    s = min(sx, sy)

    blk.XScaleFactor = s
    blk.YScaleFactor = s
    blk.ZScaleFactor = s

    # ------- IMPORTANT PART: FIX OFFSET -------
    # Find new extents after scaling
    blk.Update()
    bb2 = blk.GetBoundingBox()
    (xmin2, ymin2, zmin2) = bb2[0]

    # Offset required to move LL-corner to insert point
    dx = x - xmin2
    dy = y - ymin2

    blk.Move(make_point_variant(0, 0, 0), make_point_variant(dx, dy, 0))
    # ------------------------------------------

    return blk



# -----------------------------
# Main fixed rating plate drawer
# -----------------------------
def draw_rating_plate(doc,
                      plate_w=150.0, plate_h=110.0,
                      margin=3.0, offset_x=100.0, offset_y=100.0,
                      label_w=40.0, product_h=20.0, row_h=10.0,
                      logo_w=35.0, logo_h=20.0,
                      draw_logo_box=True):
    
    dimension_text_size = 5  # default dimension text size

    style = ensure_consolas_style(doc)
    bold_style = ensure_consolas_bold_style(doc)
    
    ms = doc.ModelSpace

    ox = offset_x
    oy = offset_y
    w = plate_w
    h = plate_h

    # Outer + inner frames
    add_rect(ms, ox, oy, ox+w, oy+h)
    add_rect(ms, ox+margin, oy+margin, ox+w-margin, oy+h-margin)

    # -------------------------
    # DIMENSIONS (WIDTH + HEIGHT)
    # -------------------------

    # Width dimension (bottom)
    add_dimension_aligned(
        ms,
        ox, oy,                                      # start point (bottom left of plate)
        ox + w, oy,                                  # end point (bottom right of plate)
        ox + w/2, oy - 8,                            # dimension line position (8 units below bottom)
        f"{w:.1f} mm",                               # override text
        text_height=dimension_text_size              # ADDED text_height
    )

    # Height dimension (left)
    add_dimension_aligned(
        ms,
        ox, oy,                                     # start point (bottom left of plate)
        ox, oy + h,                                 # end point (top left of plate)
        ox - 10, oy + h/2,                          # dimension line offset left (10 units left of plate)
        f"{h:.1f} mm",                              # override text
        text_height=dimension_text_size             # ADDED text_height
    )

    ux1 = ox + margin
    uy1 = oy + margin
    ux2 = ox + w - margin
    uy2 = oy + h - margin

    # small gap so divider doesn't overlay text
    sep_gap = 2.0

    # data_x will be computed after vertical divider position
    # but we pick label column width = label_w (user supplied)

    # start Y at top inner
    y_top_inner = uy2
    y = y_top_inner

    # -------------------------
    # PRODUCT row (MTEXT)
    # -------------------------
    y_bottom_product = y - product_h
    add_rect(ms, ux1, y_bottom_product, ux2, y)

    
    # draw label text (left column)
    add_text(ms, "PRODUCT", ux1 + 3, y - 6, 3.2, bold_style)
    

    # compute divider x (slightly right of label_w to avoid touching label)
    vertical_shift = 3.0
    vx = ux1 + label_w + sep_gap + vertical_shift
    # data start x (a bit to the right of divider)
    data_x = vx + 2.0

    # compute available width for mtext: from data_x to ux2 - small margin
    mtext_w = (ux2 - data_x) - 4.0
    # place MTEXT: position param is lower-left of mtext, so provide (data_x, top_of_box - height_offset)
    # We'll place MTEXT top at y - 2, and set MText to fit within mtext_w
    mt_top = y - 2
    product_font_h = 3.0
    add_mtext(ms,
              r"\fConsolas|b1;110V 40A FLOAT CUM BOOST BATTERY CHARGER SUITABLE FOR 110V 200AH VRLA BATTERY",
              data_x, mt_top, mtext_w, product_font_h, style)

    # move cursor below product
    y = y_bottom_product

    # -------------------------
    # RATED VOLTAGE row (2-col)
    # -------------------------
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, "RATED VOLTAGE", ux1 + 3, y - 6, 3.0, style)
    
    add_text(ms, "415V AC, 3 PHASE, 4 WIRES, 50HZ", data_x, y - 6, 3.0, style)
    y = y_bottom

    # -------------------------
    # 3-column block (OUTPUT VOLTAGE + OUTPUT CURRENT)
    # -------------------------
    three_top = y

    # OUTPUT VOLTAGE
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, "OUTPUT VOLTAGE", ux1 + 3, y - 6, 3.0, style)
    # center float column width heuristic
    col1_x = data_x
    col1_w = (ux2 - data_x) * 0.5 - 6
    col2_x = data_x + col1_w + 8  # boost start x
    add_text(ms, "FLOAT : 123.75V", col1_x + 2, y - 6, 3.0, style)
    add_text(ms, "BOOST : 126.5V", col2_x + 2, y - 6, 3.0, style)
    y = y_bottom

    # OUTPUT CURRENT
    y_bottom2 = y - row_h
    add_rect(ms, ux1, y_bottom2, ux2, y)
    add_text(ms, "OUTPUT CURRENT", ux1 + 3, y - 6, 3.0, style)
    add_text(ms, "FLOAT : 20A", col1_x + 2, y - 6, 3.0, style)
    add_text(ms, "BOOST : 20A", col2_x + 2, y - 6, 3.0, style)
    y = y_bottom2

    three_bottom = y

    # draw vertical lines for the three-column block:
    # float-boost divider (between float and boost) at position col2_x - small shift
    v2_x = col2_x - 4.0
    v2 = make_safearray_3d([(v2_x, three_top, 0),(v2_x, three_bottom, 0)])
    ms.AddPolyline(v2)

    # -------------------------
    # SL NO
    # -------------------------
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, "SL. NO.", ux1 + 3, y - 6, 3.0, style)
    add_text(ms, "LL/25-26/1077-OP2111/BCH", data_x, y - 6, 3.0, style)
    y = y_bottom

    # -------------------------
    # YEAR
    # -------------------------
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, "YEAR OF MFG.", ux1 + 3, y - 6, 3.0, style)
    add_text(ms, "2025", data_x, y - 6, 3.0, style)
    y = y_bottom

    # -------------------------
    # FIRST COLUMN VERTICAL LINE (stop before footer)
    # -------------------------
    v_line = make_safearray_3d([(vx, uy2, 0),(vx, y, 0)])
    ms.AddPolyline(v_line)

    # -------------------------
    # Footer block (company + logo) – FIXED (only 3 sides)
    # -------------------------

    FOOT_H = 25
    y_footer_top = y
    y_footer_bottom = y_footer_top - FOOT_H

    # Draw ONLY 3 sides (top, left, right)
    add_line(ms, ux1, y_footer_top, ux2, y_footer_top)     # top
    add_line(ms, ux1, y_footer_bottom, ux1, y_footer_top)  # left
    add_line(ms, ux2, y_footer_bottom, ux2, y_footer_top)  # right

    # Font sizes
    footer_title_h  = 3.2
    footer_text_h   = 2.6
    footer_text_h_a = 2.3

    fx = ux1 + 3

    # LOGO BOX
    lx1 = ux2 - logo_w + 3
    ly1 = y_footer_bottom - 2
    lx2 = lx1 + logo_w -3
    ly2 = ly1 + logo_h

    if draw_logo_box:        # <--- FLAG
        add_rect(ms, lx1, ly1, lx2, ly2)

    # Footer text
    # add_text(ms, "LIVELINE ELECTRONICS", fx, y_footer_top - 7, footer_title_h, style)

    add_mtext(
    ms,
    r"\fConsolas|b1;LIVELINE ELECTRONICS",
    fx,  # slight indent to align with bold text  
    y_footer_top - 4,
    200,             # width (large enough so it stays single line)
    footer_title_h,
    style
    )
    
    add_text(ms, "North Ramchandrapur, Narendrapur, Kolkata : 700103",
            fx, y_footer_top - 12, footer_text_h_a, style)

    add_text(ms, f"{align_label('Telefax')}",
            fx, y_footer_top - 17, footer_text_h, style)
    add_text(ms, f"{align_label(':')}",
            fx+15, y_footer_top - 17, footer_text_h, style)
    add_text(ms, f"{align_label('033 2477 2094')}",
            fx+25, y_footer_top - 17, footer_text_h, style)

    add_text(ms, f"{align_label('Email')}",
            fx, y_footer_top - 22, footer_text_h, style)
    add_text(ms, f"{align_label(':')}",
            fx+15, y_footer_top - 22, footer_text_h, style)
    add_text(ms, f"{align_label('info@livelineindia.com')}",
            fx+25, y_footer_top - 22, footer_text_h, style)


    logo_block = os.path.abspath("liveline_logo.dwg")

    insert_scaled_block(
        ms,
        logo_block,
        lx1 - 4,
        ly1 + 1,
        logo_w ,
        logo_h 
    )
    # -------------------------




    # zoom extents
    try:
        doc.SendCommand("_ZOOM _E ")
    except Exception:
        pass

    print("Done. Divider positions: vx=%.2f, float/boost divider ~%.2f" % (vx, v2_x))

# -----------------------------
# Run
# -----------------------------
if __name__ == "__main__":
    pythoncom.CoInitialize()
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
    except Exception:
        acad = win32com.client.Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument

    # call with your desired plate size if different
    draw_rating_plate(doc,
                      plate_w=150.0,
                      plate_h=103.0,
                      margin=3.0,
                      offset_x=100.0,
                      offset_y=100.0,
                      label_w=40.0,
                      product_h=20.0,
                      row_h=10.0,
                      logo_w=40.0,
                      logo_h=30.0,
                      draw_logo_box=False)

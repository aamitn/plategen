import sys
import os
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
                             QScrollArea, QMessageBox, QSpinBox, QDoubleSpinBox,
                             QListWidget, QListWidgetItem)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

APP_NAME = 'DB Rating Plate Generator'


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
    pts = [(x1, y1, 0), (x2, y1, 0), (x2, y2, 0), (x1, y2, 0), (x1, y1, 0)]
    v = make_safearray_3d(pts)
    if win32com is None:
        return None
    pl = ms.AddPolyline(v)
    pl.Closed = True
    return pl


def add_line(ms, x1, y1, x2, y2):
    if win32com is None:
        return None
    return ms.AddLine(make_point_variant(x1, y1), make_point_variant(x2, y2))


def add_text(ms, text, x, y, height):
    if win32com is None:
        return None
    t = ms.AddText(str(text), make_point_variant(x, y), float(height))
    return t


def add_mtext(ms, text, x, y, width, height):
    if win32com is None:
        return None
    mt = ms.AddMText(make_point_variant(x, y), float(width), str(text))
    mt.Height = float(height)
    try:
        mt.Attachment = 2
    except Exception:
        pass
    return mt


def draw_db_plate(doc, config):
    """Draw ACDB/DCDB rating plate based on config.
    This is a simple representation: outer frame, rows for PRODUCT/Input/INCOMER/OUTGOING rows,
    and cells for outgoings grouped 2 per cell horizontally.
    """
    plate_w = config.get('plate_width', 150.0)
    plate_h = config.get('plate_height', 105.0)
    ox = config.get('offset_x', 100.0)
    oy = config.get('offset_y', 100.0)
    margin = config.get('margin', 3.0)

    ms = doc.ModelSpace

    outer_top = oy + plate_h
    outer_bottom = oy

    # frame
    add_rect(ms, ox, outer_bottom, ox + plate_w, outer_top)
    add_rect(ms, ox + margin, outer_bottom + margin, ox + plate_w - margin, outer_top - margin)

    ux1 = ox + margin
    ux2 = ox + plate_w - margin
    y = outer_top - margin

    row_h = 10.5
    txt_h = 3.2

    # PRODUCT row
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, 'PRODUCT', ux1 + 2, y - 6, txt_h)
    add_mtext(ms, config.get('product_text', 'AC DISTRIBUTION BOARD'), ux1 + 50, y - 6, ux2 - (ux1 + 50), txt_h + 0.2)
    y = y_bottom

    # INPUT VOLTAGE row
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, 'INPUT VOLTAGE', ux1 + 2, y - 6, txt_h)
    add_text(ms, config.get('input_voltage', ''), ux1 + 50, y - 6, txt_h)
    y = y_bottom

    # INCOMER row
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, 'INCOMER', ux1 + 2, y - 6, txt_h)
    add_text(ms, config.get('incomer', ''), ux1 + 50, y - 6, txt_h)
    y = y_bottom

    # OUTGOINGS area - determine cells
    out_list = config.get('outgoings', [])
    n_out = len(out_list)
    # group 2 outgoings per cell horizontally
    per_cell = 2
    n_cells = max(1, ceil(n_out / per_cell))

    # allocate horizontal area for cells
    cell_w = (ux2 - ux1 - (n_cells - 1) * 4) / n_cells
    cell_x = ux1

    # draw outgoing header row box
    y_bottom = y - (row_h + 2)
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, 'OUTGOING', ux1 + 2, y - 6, txt_h)

    # draw each cell box inside outgoing area and list up to 2 entries
    for ci in range(n_cells):
        cx1 = cell_x + ci * (cell_w + 4)
        cx2 = cx1 + cell_w
        add_rect(ms, cx1, y_bottom, cx2, y - 0)
        # entries index
        start_idx = ci * per_cell
        for r in range(per_cell):
            idx = start_idx + r
            ey = y - 6 - r * (txt_h + 2)
            if idx < n_out:
                it = out_list[idx]
                # format: "6A 2P MCB - 27 NOS."
                rating = it.get('rating', '')
                poles = it.get('poles', '')
                btype = it.get('type', '')
                count = it.get('count', 1)
                txt = f"{rating}A {poles}P {btype} - {count} NOS."
                add_text(ms, txt, cx1 + 4, ey, txt_h)

    y = y_bottom - (row_h * 0.5)

    # SL NO and YEAR
    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, 'SL. NO.', ux1 + 2, y - 6, txt_h)
    add_text(ms, config.get('serial', ''), ux1 + 50, y - 6, txt_h)
    y = y_bottom

    y_bottom = y - row_h
    add_rect(ms, ux1, y_bottom, ux2, y)
    add_text(ms, 'YEAR OF MFG.', ux1 + 2, y - 6, txt_h)
    add_text(ms, str(config.get('year', datetime.now().year)), ux1 + 50, y - 6, txt_h)
    y = y_bottom

    # Footer + logo
    FOOT_H = 20
    y_footer_top = y
    y_footer_bottom = y_footer_top - FOOT_H
    add_line(ms, ux1, y_footer_top, ux2, y_footer_top)
    try:
        add_mtext(ms, r"\fConsolas|b1;LIVELINE ELECTRONICS", ux1 + 2, y_footer_top - 6, 160, 3.6)
    except Exception:
        add_mtext(ms, 'LIVELINE ELECTRONICS', ux1 + 2, y_footer_top - 6, 160, 3.6)

    # logo (reuse liveline_logo.dwg if present)
    logo_block = os.path.abspath('liveline_logo.dwg')
    if os.path.exists(logo_block):
        try:
            ins_pt = make_point_variant(ux2 - 46, y_footer_bottom - 6)
            blk = ms.InsertBlock(ins_pt, logo_block, 1.0, 1.0, 1.0, 0)
            try:
                blk.Update()
            except Exception:
                pass
            # attempt a regen
            try:
                doc.SendCommand('_REGEN ')
            except Exception:
                try:
                    doc.Regen(0)
                except Exception:
                    pass
        except Exception:
            pass

    try:
        doc.SendCommand('_ZOOM _E ')
    except Exception:
        pass


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
        self.resize(820, 640)

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
        btn.clicked.connect(self.generate_plate)
        L.addWidget(btn)

    def create_config_group(self):
        g = QGroupBox('General')
        l = QGridLayout()
        l.addWidget(QLabel('DB Type:'), 0, 0)
        self.db_type = QComboBox(); self.db_type.addItems(['ACDB','DCDB']); l.addWidget(self.db_type, 0, 1)

        l.addWidget(QLabel('Product suffix:'), 0, 2)
        self.product = QLineEdit('AC DISTRIBUTION BOARD'); l.addWidget(self.product, 0, 3)

        l.addWidget(QLabel('Input Voltage:'), 1, 0)
        self.input_voltage = QLineEdit('230V AC, 1 PH, 2 WIRES, 50HZ')
        l.addWidget(self.input_voltage, 1, 1, 1, 3)

        l.addWidget(QLabel('Incomer:'), 2, 0)
        self.incomer = QLineEdit('80A 3P MCCB')
        l.addWidget(self.incomer, 2, 1, 1, 3)

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
        cfg['serial'] = f"LL/{(self.year() if hasattr(self,'year') else datetime.now().year)}/{1}-OP{1}"
        cfg['year'] = datetime.now().year
        # outgoings
        outs = []
        for i in range(self.out_list_widget.count()):
            itm = self.out_list_widget.item(i)
            outs.append(itm.data(Qt.ItemDataRole.UserRole))
        cfg['outgoings'] = outs
        # plate geometry
        cfg['plate_width'] = 150.0
        cfg['plate_height'] = 105.0
        cfg['offset_x'] = 100.0
        cfg['offset_y'] = 100.0
        return cfg

    def generate_plate(self):
        cfg = self.get_config()
        if win32com is None:
            QMessageBox.information(self, 'Planned Plate', 'AutoCAD not available. Planned plate:
' + cfg['product_text'])
            return
        try:
            acad = win32com.client.Dispatch('AutoCAD.Application')
            doc = acad.ActiveDocument
        except Exception as e:
            QMessageBox.critical(self, 'AutoCAD Error', f'Could not access AutoCAD: {e}')
            return

        try:
            draw_db_plate(doc, cfg)
            QMessageBox.information(self, 'Done', 'DB plate generated in AutoCAD')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to generate plate: {e}')


def main():
    app = QApplication(sys.argv)
    w = DBRatingPlateGUI()
    w.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()

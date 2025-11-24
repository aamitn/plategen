#  ----UNDER CONSSTURCTION APP--- app_mgen_bch.py ----UNDER CONSSTURCTION APP---
import sys
from PyQt6.QtWidgets import QApplication, QMessageBox

def main():
    app = QApplication(sys.argv)

    msg = QMessageBox()
    msg.setWindowTitle("Under Construction")
    msg.setText("⚠ This application is currently under construction.\n\nFeatures will be available in a later update.")
    msg.setIcon(QMessageBox.Icon.Information)
    msg.exec()

    sys.exit()

if __name__ == "__main__":
    main()

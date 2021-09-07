import sys
from PyQt5.QtWidgets import QApplication

# GUI
from app.gui import MainWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

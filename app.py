import sys
from PySide6.QtWidgets import QApplication
from gui.main_window import JewishGenApp


def main():
    app = QApplication(sys.argv)
    window = JewishGenApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

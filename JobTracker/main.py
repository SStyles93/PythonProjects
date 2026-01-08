# main.py
import sys
from pathlib import Path
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication
from models.database import ApplicationDatabase
from controllers.controller import Controller
from views.main_window import MainWindow

# ------------------------- Helper -------------------------
def resource_path(relative_path):
    """Get absolute path to resource, works for PyInstaller."""
    try:
        # PyInstaller stores temp folder in _MEIPASS
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        base_path = Path(__file__).parent
    return base_path / relative_path

# ------------------------- Main -------------------------
def main():
    # Enable high-DPI scaling (for 4K)
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)

    # Global font size
    font = app.font()
    font.setPointSize(11)  # adjust for 4K
    app.setFont(font)

    # Load dark style
    qss_file = resource_path("styles/dark_style.qss")
    if qss_file.exists():
        with open(qss_file, "r", encoding="utf-8") as f:
            app.setStyleSheet(f.read())

    # Set global icon
    icon_file = resource_path("icons/jobat.ico")
    if icon_file.exists():
        app.setWindowIcon(QIcon(str(icon_file)))

    # === Load/create database ===  
    # Determine folder next to EXE or script
    base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
    data_file = base_path / "applications.json"
    if not data_file.exists():
        data_file.write_text("[]", encoding="utf-8")
    database = ApplicationDatabase(str(data_file))

    # Create controller
    controller = Controller(database)

    # Launch main window
    window = MainWindow(controller)
    if icon_file.exists():
        window.setWindowIcon(QIcon(str(icon_file)))
    window.show()

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

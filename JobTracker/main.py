import sys
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QApplication
from models.database import ApplicationDatabase
from controllers.controller import Controller
from views.main_window import MainWindow

def main():

    # Enable high-DPI scaling
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)

        # Global font size
    font = app.font()
    font.setPointSize(11)  # try 11–12 for 4K, adjust as needed
    app.setFont(font)

    # Dark mode style
    with open("styles/dark_style.qss", "r") as f:
        app.setStyleSheet(f.read())

    # Create database
    database = ApplicationDatabase("data/applications.json")
    controller = Controller(database)

    # Create main window
    window = MainWindow(controller)
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

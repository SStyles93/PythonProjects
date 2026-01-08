from PyQt5.QtWidgets import QMainWindow, QTableWidgetItem
from PyQt5.QtCore import QDate
from views.application_dialog import ApplicationDialog
from enums.application_status import ApplicationStatus
from ui_mainwindow import Ui_MainWindow

class MainWindow(QMainWindow):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller

        # Load the UI
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Connect buttons
        self.ui.pushButtonAdd.clicked.connect(self.on_add)
        self.ui.pushButtonEdit.clicked.connect(self.on_edit)
        self.ui.pushButtonDelete.clicked.connect(self.on_delete)
        self.ui.pushButtonSearch.clicked.connect(self.on_search)

        # Table setup
        self.ui.tableWidgetApps.setColumnCount(4)
        self.ui.tableWidgetApps.setHorizontalHeaderLabels(["Company", "Date", "Status", "Link"])
        self.ui.tableWidgetApps.setSelectionBehavior(self.ui.tableWidgetApps.SelectRows)
        self.ui.tableWidgetApps.setEditTriggers(self.ui.tableWidgetApps.NoEditTriggers)
        self.ui.tableWidgetApps.horizontalHeader().setStretchLastSection(True)

        # Status filter
        self.ui.comboBoxStatus.addItem("All")
        self.ui.comboBoxStatus.addItems(ApplicationStatus.values())

         # --- Set search date fields to today and format ---
        today = QDate.currentDate()
        self.ui.dateEditFrom.setDate(today)
        self.ui.dateEditTo.setDate(today)
        self.ui.dateEditFrom.setDisplayFormat("dd/MM/yyyy")
        self.ui.dateEditTo.setDisplayFormat("dd/MM/yyyy")

        # Initial refresh
        self.refresh_table(self.controller.database.applications)

    # Table refresh
    def refresh_table(self, applications):
        self.ui.tableWidgetApps.setRowCount(0)
        for app in applications:
            row = self.ui.tableWidgetApps.rowCount()
            self.ui.tableWidgetApps.insertRow(row)
            self.ui.tableWidgetApps.setItem(row, 0, QTableWidgetItem(app.company))
            self.ui.tableWidgetApps.setItem(row, 1, QTableWidgetItem(app.date.isoformat()))
            self.ui.tableWidgetApps.setItem(row, 2, QTableWidgetItem(app.status.value))
            self.ui.tableWidgetApps.setItem(row, 3, QTableWidgetItem(app.link))

    def get_selected_application(self):
        row = self.ui.tableWidgetApps.currentRow()
        if row < 0:
            return None
        company = self.ui.tableWidgetApps.item(row, 0).text()
        date = self.ui.tableWidgetApps.item(row, 1).text()
        for app in self.controller.database.applications:
            if app.company == company and app.date.isoformat() == date:
                return app
        return None

    # Actions
    def on_search(self):
        company = self.ui.lineEditCompany.text().strip()
        status = self.ui.comboBoxStatus.currentText()
        date_from = self.ui.dateEditFrom.date().toPyDate()
        date_to = self.ui.dateEditTo.date().toPyDate()
        results = self.controller.search_applications(company, status, date_from, date_to)
        self.refresh_table(results)

    def on_add(self):
        dialog = ApplicationDialog(self)
        if dialog.exec_():
            self.controller.create_application(dialog.get_data())
            self.refresh_table(self.controller.database.applications)

    def on_edit(self):
        app = self.get_selected_application()
        if not app:
            return
        dialog = ApplicationDialog(self, app)
        if dialog.exec_():
            self.controller.update_application(app.id, dialog.get_data())
            self.refresh_table(self.controller.database.applications)

    def on_delete(self):
        app = self.get_selected_application()
        if not app:
            return
        self.controller.delete_application(app.id)
        self.refresh_table(self.controller.database.applications)

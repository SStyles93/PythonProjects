from PyQt5.QtWidgets import QMainWindow, QTableWidgetItem
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QIcon
from views.application_dialog import ApplicationDialog
from enums.application_status import ApplicationStatus
from ui_mainwindow import Ui_MainWindow

class MainWindow(QMainWindow):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.setWindowIcon(QIcon("icon_JobAT.ico"))

        # Load the UI
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Connect buttons
        self.ui.pushButtonAdd.clicked.connect(self.on_add)
        self.ui.pushButtonEdit.clicked.connect(self.on_edit)
        self.ui.pushButtonDelete.clicked.connect(self.on_delete)
        self.ui.pushButtonSearch.clicked.connect(self.on_search)

        # Table setup
        self.ui.tableWidgetApps.setColumnCount(6)
        self.ui.tableWidgetApps.setHorizontalHeaderLabels([
            "Company", "Job Name", "Date", "Status", "Link", "Comment"
            ])
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

    def _make_item(self, text):
        return QTableWidgetItem(str(text))

    # Table refresh
    def refresh_table(self, applications):
        self.ui.tableWidgetApps.setRowCount(0)
        for app in applications:
            row = self.ui.tableWidgetApps.rowCount()
            self.ui.tableWidgetApps.insertRow(row)
            self.ui.tableWidgetApps.setItem(row, 0, self._make_item(app.company))
            self.ui.tableWidgetApps.setItem(row, 1, self._make_item(app.job_name))
            self.ui.tableWidgetApps.setItem(row, 2, self._make_item(app.date.strftime("%d/%m/%Y")))
            self.ui.tableWidgetApps.setItem(row, 3, self._make_item(app.status.value))
            self.ui.tableWidgetApps.setItem(row, 4, self._make_item(app.link))
            self.ui.tableWidgetApps.setItem(row, 5, self._make_item(app.comment))

    def get_selected_application(self):
        row = self.ui.tableWidgetApps.currentRow()
        if row < 0:
            return None
        company = self.ui.tableWidgetApps.item(row, 0).text()
        date_str = self.ui.tableWidgetApps.item(row, 2).text()
        for app in self.controller.database.applications:
            if app.company == company and app.date.strftime("%d/%m/%Y") == date_str:
                return app
        return None

    # Actions
    def on_search(self):
        company = self.ui.lineEditCompany.text().strip()
        job_name = self.ui.lineEditJobName.text().strip()
        status = self.ui.comboBoxStatus.currentText()
        date_from = self.ui.dateEditFrom.date().toPyDate()
        date_to = self.ui.dateEditTo.date().toPyDate()

        results = self.controller.search_applications(
            company=company or None,
            job_name=job_name or None,
            status=status,
            date_from=date_from,
            date_to=date_to
        )
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

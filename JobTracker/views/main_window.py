from PyQt5.QtWidgets import QMainWindow, QTableWidgetItem
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QIcon
from views.application_dialog import ApplicationDialog
from enums.application_status import ApplicationStatus
from ui.ui_mainwindow import Ui_MainWindow


class MainWindow(QMainWindow):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.setWindowIcon(QIcon("icons/jobat.ico"))

        # Load the UI
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Connect buttons
        self.ui.pushButtonAdd.clicked.connect(self.on_add)
        self.ui.pushButtonEdit.clicked.connect(self.on_edit)
        self.ui.pushButtonDelete.clicked.connect(self.on_delete)
        self.ui.pushButtonSearch.clicked.connect(self.on_search)

        # Table setup
        table = self.ui.tableWidgetApps
        table.setColumnCount(6)
        table.setHorizontalHeaderLabels([
            "Company", "Job Name", "Date", "Status", "Link", "Comment"
        ])
        table.setSelectionBehavior(table.SelectRows)
        table.setEditTriggers(table.NoEditTriggers)
        table.horizontalHeader().setStretchLastSection(True)

        # Enable sorting
        header = table.horizontalHeader()
        table.setSortingEnabled(True)
        header.setSortIndicatorShown(True)
        header.setSectionsClickable(True)
        header.sectionClicked.connect(self.on_header_clicked)  # intercept Comment clicks

        # Status filter
        self.ui.comboBoxStatus.addItem("All")
        self.ui.comboBoxStatus.addItems(ApplicationStatus.values())

        # Set search date fields to today
        today = QDate.currentDate()
        self.ui.dateEditFrom.setDate(today)
        self.ui.dateEditTo.setDate(today)
        self.ui.dateEditFrom.setDisplayFormat("dd/MM/yyyy")
        self.ui.dateEditTo.setDisplayFormat("dd/MM/yyyy")

        # Initial refresh
        self.refresh_table(self.controller.database.applications)

        # Default sort by Date descending
        table.sortItems(2, Qt.DescendingOrder)
        header.setSortIndicator(2, Qt.DescendingOrder)

    # ----------------------------
    # Ignore Comment column clicks
    # ----------------------------
    def on_header_clicked(self, column):
        if column == 5:  # Comment column
            return
        # For all other columns, Qt handles sorting and arrow automatically

    # ----------------------------
    # Helper: create QTableWidgetItem
    # ----------------------------
    def _make_item(self, text, sort_data=None):
        item = QTableWidgetItem(str(text))
        # If no sort_data provided, use displayed text (Qt sorts by text by default)
        if sort_data is None:
            sort_data = str(text)
        item.setData(Qt.UserRole, sort_data)
        return item

    # ----------------------------
    # Refresh table (view-only sort)
    # ----------------------------
    def refresh_table(self, applications):
        table = self.ui.tableWidgetApps

        # Disable sorting while filling to avoid flicker
        sorting_enabled = table.isSortingEnabled()
        table.setSortingEnabled(False)

        table.setRowCount(0)
        for app in applications:
            row = table.rowCount()
            table.insertRow(row)

            table.setItem(row, 0, self._make_item(app.company))
            table.setItem(row, 1, self._make_item(app.job_name))
            table.setItem(row, 2, DateItem(app.date))   # <-- use custom DateItem
            table.setItem(row, 3, self._make_item(app.status.value))
            table.setItem(row, 4, self._make_item(app.link))
            table.setItem(row, 5, self._make_item(app.comment))

        # Restore sorting
        table.setSortingEnabled(sorting_enabled)

        # Set arrow after refresh
        table.horizontalHeader().setSortIndicator(2, Qt.DescendingOrder)


    # ----------------------------
    # Selected application
    # ----------------------------
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

    # ----------------------------
    # Actions
    # ----------------------------
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


# Custom Table Item to sort by date
class DateItem(QTableWidgetItem):
    def __init__(self, date):
        super().__init__(date.strftime("%d/%m/%Y"))
        self.date = date

    def __lt__(self, other):
        if isinstance(other, DateItem):
            return self.date < other.date
        return super().__lt__(other)
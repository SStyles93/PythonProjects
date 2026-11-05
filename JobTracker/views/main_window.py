from PyQt5.QtWidgets import QMainWindow, QTreeWidgetItem
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QIcon, QFont
from views.application_dialog import ApplicationDialog
from enums.application_status import ApplicationStatus
from ui.ui_mainwindow import Ui_MainWindow


DATE_COLUMN = 2


class MainWindow(QMainWindow):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.setWindowIcon(QIcon("icons/jobat.ico"))

        # Months the user has collapsed (kept across refreshes)
        self.collapsed_months = set()

        # Load the UI
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Connect buttons
        self.ui.pushButtonAdd.clicked.connect(self.on_add)
        self.ui.pushButtonEdit.clicked.connect(self.on_edit)
        self.ui.pushButtonDelete.clicked.connect(self.on_delete)
        self.ui.pushButtonSearch.clicked.connect(self.on_search)

        # Tree setup (one collapsible group per month)
        tree = self.ui.treeWidgetApps
        tree.setColumnCount(6)
        tree.setHeaderLabels([
            "Company", "Job Name", "Date", "Status", "Link", "Comment"
        ])
        tree.setSelectionBehavior(tree.SelectRows)
        tree.setEditTriggers(tree.NoEditTriggers)
        tree.header().setStretchLastSection(True)
        tree.itemExpanded.connect(self.on_month_expanded)
        tree.itemCollapsed.connect(self.on_month_collapsed)
        tree.itemDoubleClicked.connect(self.on_item_double_clicked)

        # Enable sorting
        header = tree.header()
        tree.setSortingEnabled(True)
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

        # Default sort by Date descending
        tree.sortItems(DATE_COLUMN, Qt.DescendingOrder)

        # Initial refresh
        self.refresh_table(self.controller.database.applications)

    # ----------------------------
    # Ignore Comment column clicks
    # ----------------------------
    def on_header_clicked(self, column):
        if column == 5:  # Comment column
            return
        # For all other columns, Qt handles sorting and arrow automatically

    # ----------------------------
    # Remember collapsed months
    # ----------------------------
    def on_month_expanded(self, item):
        if isinstance(item, MonthItem):
            self.collapsed_months.discard(item.month)

    def on_month_collapsed(self, item):
        if isinstance(item, MonthItem):
            self.collapsed_months.add(item.month)

    def on_item_double_clicked(self, item, column):
        if isinstance(item, ApplicationItem):
            self.on_edit()

    # ----------------------------
    # Refresh tree (grouped by month, view-only sort)
    # ----------------------------
    def refresh_table(self, applications):
        tree = self.ui.treeWidgetApps

        # Disable sorting while filling to avoid flicker
        sorting_enabled = tree.isSortingEnabled()
        tree.setSortingEnabled(False)
        tree.clear()

        months = {}
        for app in applications:
            month = app.date.replace(day=1)
            months.setdefault(month, []).append(app)

        for month, apps in months.items():
            month_item = MonthItem(month, len(apps))
            tree.addTopLevelItem(month_item)
            month_item.setFirstColumnSpanned(True)
            for app in apps:
                month_item.addChild(ApplicationItem(app))
            month_item.setExpanded(month not in self.collapsed_months)

        # Restore sorting (re-sorts by the current header column/order)
        tree.setSortingEnabled(sorting_enabled)

    # ----------------------------
    # Selected application
    # ----------------------------
    def get_selected_application(self):
        item = self.ui.treeWidgetApps.currentItem()
        if not isinstance(item, ApplicationItem):
            return None

        for app in self.controller.database.applications:
            if app.id == item.app_id:
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


# Collapsible month header row
class MonthItem(QTreeWidgetItem):
    def __init__(self, month, count):
        super().__init__([f"{month.strftime('%B %Y')} ({count})"])
        self.month = month
        font = QFont()
        font.setBold(True)
        self.setFont(0, font)
        self.setFlags(Qt.ItemIsEnabled)  # not selectable

    def __lt__(self, other):
        if not isinstance(other, MonthItem):
            return super().__lt__(other)
        tree = self.treeWidget()
        if tree.sortColumn() == DATE_COLUMN:
            return self.month < other.month
        # Sorting by another column: keep months newest first
        if tree.header().sortIndicatorOrder() == Qt.AscendingOrder:
            return self.month > other.month
        return self.month < other.month


# Application row, sorts by real date in the Date column
class ApplicationItem(QTreeWidgetItem):
    def __init__(self, app):
        super().__init__([
            app.company,
            app.job_name,
            app.date.strftime("%d/%m/%Y"),
            app.status.value,
            app.link,
            app.comment,
        ])
        self.app_id = app.id
        self.date = app.date

    def __lt__(self, other):
        if isinstance(other, ApplicationItem) and self.treeWidget().sortColumn() == DATE_COLUMN:
            return self.date < other.date
        return super().__lt__(other)

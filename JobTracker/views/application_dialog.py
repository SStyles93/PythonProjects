from PyQt5.QtWidgets import QDialog
from ui_application_dialog import Ui_ApplicationDialog
from enums.application_status import ApplicationStatus
from PyQt5.QtCore import QDate

class ApplicationDialog(QDialog, Ui_ApplicationDialog):
    def __init__(self, parent=None, application=None):
        super().__init__(parent)
        self.setupUi(self)
        self.application = application

        # Fill status combo
        self.status_input.addItems(ApplicationStatus.values())

        # Date format
        self.date_input.setDisplayFormat("dd/MM/yyyy")

        # Set date to today if new application
        if self.application is None:
            self.date_input.setDate(QDate.currentDate())
        else:
            self.date_input.setDate(self.application.date)

        # Fill data if editing
        if self.application:
            self.company_input.setText(self.application.company)
            self.jobName_input.setText(self.application.job_name)
            self.date_input.setDate(self.application.date)
            self.link_input.setText(self.application.link)
            self.status_input.setCurrentText(self.application.status.value)
            self.comment_input.setPlainText(self.application.comment)

        # Connect buttons
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)

    def get_data(self):
        return {
            "company": self.company_input.text().strip(),
            "job_name": self.jobName_input.text().strip(),
            "date": self.date_input.date().toPyDate(),
            "link": self.link_input.text().strip(),
            "status": ApplicationStatus(self.status_input.currentText()),
            "comment": self.comment_input.toPlainText().strip()
}

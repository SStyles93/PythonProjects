import uuid
from models.database import Application
from enums.application_status import ApplicationStatus

class Controller:
    def __init__(self, database):
        self.database = database

    def create_application(self, data):
        app = Application(
            id_=str(uuid.uuid4()),
            company=data["company"],
            date_=data["date"],
            link=data["link"],
            status=data["status"]
        )
        self.database.add(app)

    def update_application(self, app_id, data):
        self.database.update(app_id, data)

    def delete_application(self, app_id):
        self.database.delete(app_id)

    def search_applications(self, company=None, status=None, date_from=None, date_to=None):
        results = self.database.applications
        if company:
            results = [app for app in results if company.lower() in app.company.lower()]
        if status and status != "All":
            results = [app for app in results if app.status.value == status]
        if date_from:
            results = [app for app in results if app.date >= date_from]
        if date_to:
            results = [app for app in results if app.date <= date_to]
        return results

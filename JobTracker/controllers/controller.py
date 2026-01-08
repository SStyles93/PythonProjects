import uuid
from models.database import Application
from enums.application_status import ApplicationStatus

class Controller:
    def __init__(self, database):
        self.database = database

     # ---------------- CRUD ----------------
    def create_application(self, data):
        from models.database import Application
        app = Application(
            company=data["company"],
            job_name=data.get("job_name", ""),
            date=data["date"],
            status=data["status"],
            link=data["link"],
            comment=data.get("comment", "")
        )
        self.database.add_application(app)

    def update_application(self, app_id, data):
        self.database.update(app_id, data)

    def delete_application(self, app_id):
        self.database.delete(app_id)

    # ---------------- SEARCH ----------------
    def search_applications(self, company=None, job_name=None, status=None, date_from=None, date_to=None):
        results = self.database.applications

        if company:
            results = [app for app in results if company.lower() in app.company.lower()]

        if job_name:
            results = [app for app in results if job_name.lower() in app.job_name.lower()]

        if status and status != "All":
            results = [app for app in results if app.status.value == status]

        if date_from:
            results = [app for app in results if app.date >= date_from]

        if date_to:
            results = [app for app in results if app.date <= date_to]

        return results

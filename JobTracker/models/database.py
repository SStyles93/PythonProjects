import json
from datetime import date
from pathlib import Path
from enums.application_status import ApplicationStatus

class Application:
    def __init__(self, id_, company, date_, link, status):
        self.id = id_
        self.company = company
        self.date = date_
        self.link = link
        self.status = status

    def to_dict(self):
        return {
            "id": self.id,
            "company": self.company,
            "date": self.date.isoformat(),
            "link": self.link,
            "status": self.status.value
        }

class ApplicationDatabase:
    def __init__(self, path):
        self.path = Path(path)
        self.applications = []
        self.load()

    def load(self):
        if self.path.exists():
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.applications = [
                    Application(
                        id_=item["id"],
                        company=item["company"],
                        date_=date.fromisoformat(item["date"]),
                        link=item["link"],
                        status=ApplicationStatus(item["status"])
                    )
                    for item in data
                ]
            except json.JSONDecodeError:
                self.applications = []
        else:
            self.applications = []

    def save(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump([app.to_dict() for app in self.applications], f, indent=4)

    # --- CRUD Operations ---
    def add(self, app):
        self.applications.append(app)
        self.save()

    def update(self, app_id, new_data):
        for i, app in enumerate(self.applications):
            if app.id == app_id:
                self.applications[i].company = new_data["company"]
                self.applications[i].date = new_data["date"]
                self.applications[i].link = new_data["link"]
                self.applications[i].status = new_data["status"]
                self.save()
                break

    def delete(self, app_id):
        self.applications = [app for app in self.applications if app.id != app_id]
        self.save()

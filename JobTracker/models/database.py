import json
import uuid
from datetime import date, datetime
from pathlib import Path
from enums.application_status import ApplicationStatus

class Application:
    def __init__(self, company, date, status, link, job_name="", comment="", id=None):
        self.id = id or str(uuid.uuid4())
        self.company = company
        self.date = date
        self.status = status
        self.link = link
        self.job_name = job_name
        self.comment = comment

    def to_dict(self):
        return {
            "id": self.id,
            "company": self.company,
            "date": self.date.isoformat(),
            "status": self.status.value,
            "link": self.link,
            "job_name": self.job_name,
            "comment": self.comment
        }

    @classmethod
    def from_dict(cls, data):
        return cls(
            id=data.get("id"),
            company=data.get("company", ""),
            date=datetime.fromisoformat(data["date"]).date(),
            status=ApplicationStatus(data.get("status")),
            link=data.get("link", ""),
            job_name=data.get("job_name", ""),
            comment=data.get("comment", "")
        )

class ApplicationDatabase:
    def __init__(self, path):
        self.path = Path(path)
        self.applications = []
        self.load()

    def load(self):
        """Load applications from JSON file."""
        if self.path.exists():
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.applications = [Application.from_dict(item) for item in data]
            except json.JSONDecodeError:
                self.applications = []
        else:
            self.applications = []

    def save(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump([app.to_dict() for app in self.applications], f, indent=4)

    # --- CRUD Operations ---
    def add_application(self, app):
        self.applications.append(app)
        self.save()

    def update(self, app_id, new_data):
        for i, app in enumerate(self.applications):
            if app.id == app_id:
                self.applications[i].company = new_data["company"]
                self.applications[i].job_name = new_data["job_name"]
                self.applications[i].date = new_data["date"]
                self.applications[i].status = new_data["status"]
                self.applications[i].link = new_data["link"]
                self.applications[i].comment = new_data["comment"]
                self.save()
                break


    def delete(self, app_id):
        self.applications = [app for app in self.applications if app.id != app_id]
        self.save()

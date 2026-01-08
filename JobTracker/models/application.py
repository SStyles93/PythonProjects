from utils.date_utils import parse_date
from enums.application_status import ApplicationStatus


class JobApplication:
    def __init__(self, app_id, company, date, link, status, notes=""):
        self.id = app_id
        self.company = company
        self.date = date
        self.link = link
        self.status = status
        self.notes = notes

    def to_dict(self):
        return {
            "id": self.id,
            "company": self.company,
            "date": self.date.isoformat(),
            "link": self.link,
            "status": self.status.value,
            "notes": self.notes
        }

    @staticmethod
    def from_dict(data):
        return JobApplication(
            app_id=data["id"],
            company=data["company"],
            date=parse_date(data["date"]),
            link=data["link"],
            status=ApplicationStatus(data["status"]),
            notes=data.get("notes", "")
        )

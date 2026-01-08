from enum import Enum

class ApplicationStatus(Enum):
    Pending = "Pending"
    Interview = "Interview"
    Contract = "Contract"
    NegativeReply = "Negative Reply"

    @classmethod
    def values(cls):
        return [s.value for s in cls]

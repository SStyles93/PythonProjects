from datetime import date


def parse_date(date_str: str) -> date:
    return date.fromisoformat(date_str)
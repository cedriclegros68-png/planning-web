import re
from datetime import datetime, timedelta
from docx import Document

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build


SCOPES = ["https://www.googleapis.com/auth/calendar"]
SERVICE_ACCOUNT_FILE = "service_account.json"


def get_service():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES
    )
    return build("calendar", "v3", credentials=creds)


def process_docx(docx_path: str):
    service = get_service()
    doc = Document(docx_path)

    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) < 6:
                continue

            date_time = cells[0]
            prestation = cells[3]
            contact = cells[4]
            address = cells[5]

            match = re.search(
                r"(\d{2}/\d{2}/\d{2}).*?(\d{2}:\d{2})",
                date_time,
                re.S
            )
            if not match:
                continue

            date_str, time_str = match.groups()
            start = datetime.strptime(f"{date_str} {time_str}", "%d/%m/%y %H:%M")
            end = start + timedelta(hours=1)

            event = {
                "summary": prestation,
                "location": address.replace("\n", " "),
                "description": f"{contact}\n\n{address}",
                "start": {"dateTime": start.isoformat(), "timeZone": "Europe/Paris"},
                "end": {"dateTime": end.isoformat(), "timeZone": "Europe/Paris"},
            }

            service.events().insert(
                calendarId="primary",
                body=event
            ).execute()

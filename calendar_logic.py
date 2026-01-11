import os
import re
from datetime import datetime, timedelta
from docx import Document
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials

SCOPES = ["https://www.googleapis.com/auth/calendar"]
TIMEZONE = "Europe/Paris"

def get_service():
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    return build("calendar", "v3", credentials=creds)

def detect_icons(text):
    icons = []
    t = text.lower()

    if "photos" in t:
        icons.append("🔴")

    if (
        ("vp18" in t or "el-vp18" in t or " 18 " in t)
        and "kva" not in t
    ):
        icons.append("🟠")

    if "qd" in t:
        icons.append("🟣")

    return " ".join(icons)

def process_docx(path):
    service = get_service()
    filename = os.path.basename(path)

    # supprimer anciens événements
    events = service.events().list(
        calendarId="primary",
        q=f"SOURCE_DOCX={filename}"
    ).execute()

    for ev in events.get("items", []):
        service.events().delete(
            calendarId="primary",
            eventId=ev["id"]
        ).execute()

    doc = Document(path)
    table = doc.tables[1]

    for i, row in enumerate(table.rows):
        if i == 0:
            continue

        cells = [c.text.strip() for c in row.cells]
        full_text = " ".join(cells)

        m_date = re.search(r"(\d{2}/\d{2}/\d{2})", cells[0])
        m_time = re.search(r"a\s*(\d{2}:\d{2})", cells[0])
        if not m_date or not m_time:
            continue

        date = datetime.strptime(m_date.group(1), "%d/%m/%y").date()
        time_ = datetime.strptime(m_time.group(1), "%H:%M").time()

        try:
            duration = timedelta(hours=float(cells[2].replace(",", ".")))
        except:
            continue

        start = datetime.combine(date, time_)
        end = start + duration

        prestation = cells[3].replace("\n", " / ")
        company = cells[5].splitlines()[0]
        icons = detect_icons(full_text)

        event = {
            "summary": f"{icons} {company} – {prestation}".strip(),
            "description": f"SOURCE_DOCX={filename}",
            "start": {"dateTime": start.isoformat(), "timeZone": TIMEZONE},
            "end": {"dateTime": end.isoformat(), "timeZone": TIMEZONE},
        }

        service.events().insert(
            calendarId="primary",
            body=event
        ).execute()

import os
import re
from datetime import datetime, timedelta

from docx import Document

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


# ==========================
# CONFIG
# ==========================

SCOPES = ["https://www.googleapis.com/auth/calendar"]


# ==========================
# GOOGLE CALENDAR SERVICE
# ==========================

def get_service():
    creds = None

    # 1️⃣ Charger token.json s'il existe
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # 2️⃣ Si pas de credentials valides → OAuth
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json",
            SCOPES
        )

        # ⚠️ Compatible serveurs cloud (Render)
        creds = flow.run_local_server(
            host="0.0.0.0",
            port=8080,
            open_browser=False
        )

        # 3️⃣ Sauvegarder token.json
        with open("token.json", "w", encoding="utf-8") as token:
            token.write(creds.to_json())

    return build("calendar", "v3", credentials=creds)


# ==========================
# DOCX → GOOGLE CALENDAR
# ==========================

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

            # --------------------------
            # DATE & HEURE
            # --------------------------
            match = re.search(
                r"(\d{2}/\d{2}/\d{2}).*?(\d{2}:\d{2})",
                date_time,
                re.S
            )
            if not match:
                continue

            date_str, time_str = match.groups()

            start = datetime.strptime(
                f"{date_str} {time_str}",
                "%d/%m/%y %H:%M"
            )
            end = start + timedelta(hours=1)

            # --------------------------
            # TITRE
            # --------------------------
            summary = prestation

            # --------------------------
            # DESCRIPTION
            # --------------------------
            description = f"{contact}\n\n{address}"

            # --------------------------
            # EVENT
            # --------------------------
            event = {
                "summary": summary,
                "location": address.replace("\n", " "),
                "description": description,
                "start": {
                    "dateTime": start.isoformat(),
                    "timeZone": "Europe/Paris",
                },
                "end": {
                    "dateTime": end.isoformat(),
                    "timeZone": "Europe/Paris",
                },
            }

            service.events().insert(
                calendarId="primary",
                body=event
            ).execute()

import requests
from openpyxl import Workbook
from datetime import datetime
from zoneinfo import ZoneInfo
import os
from dotenv import load_dotenv
import time

# -------------------------------------------------
# Load environment
# -------------------------------------------------
load_dotenv()

SPIKE_API_KEY = os.getenv("SPIKE_API_KEY")

teams = {
    key.replace("TEAM_", ""): value
    for key, value in os.environ.items()
    if key.startswith("TEAM_")
}

# -------------------------------------------------
# Time helpers
# -------------------------------------------------
def utc_to_ist(utc_time_str):
    if not utc_time_str:
        return None
    try:
        dt = datetime.fromisoformat(utc_time_str.replace("Z", "+00:00"))
        return dt.astimezone(ZoneInfo("Asia/Kolkata"))
    except:
        return None

def ist_str(dt):
    return dt.strftime("%Y-%m-%d %H:%M:%S") if dt else ""

# -------------------------------------------------
# User resolver
# -------------------------------------------------
user_cache = {}

def resolve_user(user_field, team_id):
    if not user_field:
        return ""

    if isinstance(user_field, dict):
        return (
            f"{user_field.get('firstName','')} "
            f"{user_field.get('lastName','')}"
        ).strip() or user_field.get("email", "")

    uid = str(user_field)
    if uid in user_cache:
        return user_cache[uid]

    r = requests.get(
        f"https://api.spike.sh/users/{uid}",
        headers={
            "x-api-key": SPIKE_API_KEY,
            "x-team-id": team_id,
            "Accept": "application/json"
        }
    )

    name = uid
    if r.status_code == 200:
        u = r.json()
        name = (
            f"{u.get('firstName','')} "
            f"{u.get('lastName','')}"
        ).strip() or u.get("email", uid)

    user_cache[uid] = name
    time.sleep(0.1)
    return name

# -------------------------------------------------
# FETCH INCIDENTS (USED BY UI)
# -------------------------------------------------
def fetch_incidents_for_range(from_date, to_date):
    rows = []

    for team_name, team_id in teams.items():
        headers = {
            "x-api-key": SPIKE_API_KEY,
            "x-team-id": team_id,
            "Accept": "application/json"
        }

        resp = requests.get("https://api.spike.sh/incidents", headers=headers)
        if resp.status_code != 200:
            continue

        incidents = resp.json().get("incidents", resp.json())

        for inc in incidents:
            nack_dt = utc_to_ist(inc.get("NACK_at"))
            if not nack_dt:
                continue

            if not (from_date <= nack_dt.date() <= to_date):
                continue

            grouped = inc.get("groupedIncident", {})

            notes = []
            for note in grouped.get("notes", []):
                note_dt = utc_to_ist(note.get("createdAt"))
                if note_dt and note_dt.date() == nack_dt.date():
                    user = resolve_user(note.get("user"), team_id)
                    content = note.get("content", "").replace("\n", " ")
                    notes.append(f"{ist_str(note_dt)} | {user}: {content}")

            rows.append({
                "Team Name": team_name,
                "Counter ID": inc.get("counterId"),
                "Message": inc.get("message"),
                "Assignee Email": ", ".join(
                    a.get("email", "") for a in inc.get("assignee", [])
                ),
                "Priority": inc.get("metadata", {}).get("priority"),
                "Status": inc.get("status"),
                "Source": inc.get("integration", {}).get("name"),
                "Created (IST)": ist_str(nack_dt),
                "ACK At (IST)": ist_str(utc_to_ist(inc.get("ACK_at"))),
                "Resolved At (IST)": ist_str(utc_to_ist(inc.get("RES_at"))),
                "Notes": "\n".join(notes)
            })

    return rows

# -------------------------------------------------
# GENERATE EXCEL (USED BY UI)
# -------------------------------------------------
def generate_excel(rows, from_date, to_date):
    wb = Workbook()
    ws = wb.active
    ws.title = "Spike Incidents"

    ws.append(list(rows[0].keys()))   # FIXED

    for r in rows:
        ws.append(list(r.values()))

    file_name = f"spike_incidents_{from_date}_to_{to_date}.xlsx"
    wb.save(file_name)

    return file_name

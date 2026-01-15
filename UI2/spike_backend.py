import requests
from openpyxl import Workbook
from datetime import datetime
from zoneinfo import ZoneInfo
import os
from dotenv import load_dotenv
import time

load_dotenv()

SPIKE_API_KEY = os.getenv("SPIKE_API_KEY")

teams = {
    key.replace("TEAM_", ""): value
    for key, value in os.environ.items()
    if key.startswith("TEAM_")
}

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
            "x-team-id": team_id
        }
    )

    name = uid
    if r.status_code == 200:
        u = r.json()
        name = f"{u.get('firstName','')} {u.get('lastName','')}".strip()

    user_cache[uid] = name
    time.sleep(0.1)
    return name

def fetch_incidents_for_range(from_date, to_date):
    rows = []

    for team_name, team_id in teams.items():
        resp = requests.get(
            "https://api.spike.sh/incidents",
            headers={
                "x-api-key": SPIKE_API_KEY,
                "x-team-id": team_id
            }
        )

        if resp.status_code != 200:
            continue

        for inc in resp.json().get("incidents", []):
            nack_dt = utc_to_ist(inc.get("NACK_at"))
            if not nack_dt or not (from_date <= nack_dt.date() <= to_date):
                continue

            notes = []
            for note in inc.get("groupedIncident", {}).get("notes", []):
                note_dt = utc_to_ist(note.get("createdAt"))
                if note_dt:
                    notes.append(
                        f"{ist_str(note_dt)} | "
                        f"{resolve_user(note.get('user'), team_id)}: "
                        f"{note.get('content','')}"
                    )

            rows.append({
                "Team Name": team_name,
                "Counter ID": inc.get("counterId"),
                "Message": inc.get("message"),
                "Priority": inc.get("metadata", {}).get("priority"),
                "Status": inc.get("status"),
                "Created (IST)": ist_str(nack_dt),
                "Resolved At (IST)": ist_str(utc_to_ist(inc.get("RES_at"))),
                "Notes": "\n".join(notes)
            })

    return rows

def generate_excel(rows, from_date, to_date):
    wb = Workbook()
    ws = wb.active
    ws.append(list(rows[0].keys()))
    for r in rows:
        ws.append(list(r.values()))

    file = f"spike_incidents_{from_date}_to_{to_date}.xlsx"
    wb.save(file)
    return file

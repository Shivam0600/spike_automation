import requests
from datetime import datetime
from zoneinfo import ZoneInfo
import os
import time
from openpyxl import Workbook
from dotenv import load_dotenv

# -------------------------------------------------
# ENV
# -------------------------------------------------
load_dotenv()

SPIKE_API_KEY = os.getenv("SPIKE_API_KEY")
teams = {k.replace("TEAM_", ""): v for k, v in os.environ.items() if k.startswith("TEAM_")}

IST = ZoneInfo("Asia/Kolkata")

if not SPIKE_API_KEY or not teams:
    raise RuntimeError("SPIKE_API_KEY or TEAM_* missing")

# -------------------------------------------------
# HELPERS
# -------------------------------------------------
def utc_to_ist(utc_str):
    if not utc_str:
        return None
    return datetime.fromisoformat(utc_str.replace("Z", "+00:00")).astimezone(IST)

def ist_str(dt):
    return dt.strftime("%Y-%m-%d %H:%M:%S") if dt else ""

user_cache = {}

def resolve_user(user_field, team_id):
    if not user_field:
        return ""
    if isinstance(user_field, dict):
        return f"{user_field.get('firstName','')} {user_field.get('lastName','')}".strip()

    uid = str(user_field)
    if uid in user_cache:
        return user_cache[uid]

    r = requests.get(
        f"https://api.spike.sh/users/{uid}",
        headers={"x-api-key": SPIKE_API_KEY, "x-team-id": team_id}
    )

    name = uid
    if r.status_code == 200:
        u = r.json()
        name = f"{u.get('firstName','')} {u.get('lastName','')}".strip() or u.get("email", uid)

    user_cache[uid] = name
    time.sleep(0.1)
    return name

# -------------------------------------------------
# FETCH INCIDENTS
# -------------------------------------------------
def fetch_incidents_for_range(from_dt, to_dt):
    rows = []

    for team_name, team_id in teams.items():
        headers = {"x-api-key": SPIKE_API_KEY, "x-team-id": team_id}
        resp = requests.get("https://api.spike.sh/incidents", headers=headers)
        if resp.status_code != 200:
            continue

        incidents = resp.json().get("incidents", resp.json())

        for inc in incidents:
            nack_dt = utc_to_ist(inc.get("NACK_at"))
            if not nack_dt or not (from_dt <= nack_dt <= to_dt):
                continue

            grouped = inc.get("groupedIncident", {})
            notes = []

            for note in grouped.get("notes", []):
                note_dt = utc_to_ist(note.get("createdAt"))
                if note_dt:
                    user = resolve_user(note.get("user"), team_id)
                    notes.append(f"{ist_str(note_dt)} | {user}: {note.get('content','').replace(chr(10),' ')}")

            rows.append({
                "Team Name": team_name,
                "Counter ID": inc.get("counterId"),
                "Message": inc.get("message"),
                "Assignee Email": ", ".join(a.get("email","") for a in inc.get("assignee", [])),
                "Priority": inc.get("metadata", {}).get("priority"),
                "Status": inc.get("status"),
                "Source": inc.get("integration", {}).get("name"),
                "Created (IST)": ist_str(nack_dt),
                "ACK At (IST)": ist_str(utc_to_ist(inc.get("ACK_at"))),
                "Notes": "\n".join(notes)
            })

    # 🔥 SORT BY CREATED TIME (LATEST FIRST)
    rows.sort(
        key=lambda x: datetime.strptime(x["Created (IST)"], "%Y-%m-%d %H:%M:%S"),
        reverse=True
    )

    return rows

# -------------------------------------------------
# EXCEL
# -------------------------------------------------
def generate_excel(rows, from_dt, to_dt):
    wb = Workbook()
    ws = wb.active
    ws.title = "Incident Report"

    ws.append(list(rows[0].keys()))
    for r in rows:
        ws.append(list(r.values()))

    file_name = f"spike_incidents_{from_dt.date()}_to_{to_dt.date()}.xlsx"
    wb.save(file_name)
    return file_name

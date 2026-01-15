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
if not SPIKE_API_KEY:
    print("❌ Missing SPIKE_API_KEY in .env")
    exit(1)

teams = {
    key.replace("TEAM_", ""): value
    for key, value in os.environ.items()
    if key.startswith("TEAM_")
}
if not teams:
    print("❌ No TEAM_* variables found in .env")
    exit(1)

# -------------------------------------------------
# Convert UTC → IST full datetime
# -------------------------------------------------
def utc_to_ist(utc_time_str):
    if not utc_time_str:
        return ""
    try:
        dt = datetime.fromisoformat(utc_time_str.replace("Z", "+00:00"))
        ist = dt.astimezone(ZoneInfo("Asia/Kolkata"))
        return ist
    except:
        return None

def ist_str(dt):
    if isinstance(dt, datetime):
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    return ""

# -------------------------------------------------
# Get display name for user
# -------------------------------------------------
user_cache = {}
def resolve_user(user_field, team_id):
    if not user_field:
        return ""
    if isinstance(user_field, dict):
        first = user_field.get("firstName", "")
        last = user_field.get("lastName", "")
        email = user_field.get("email", "")
        name = f"{first} {last}".strip() or email or ""
        return name

    user_id = str(user_field)
    if user_id in user_cache:
        return user_cache[user_id]

    resp = requests.get(
        f"https://api.spike.sh/users/{user_id}",
        headers={
            "x-api-key": SPIKE_API_KEY,
            "x-team-id": team_id,
            "Accept": "application/json"
        }
    )

    if resp.status_code == 200:
        udata = resp.json()
        first = udata.get("firstName", "")
        last = udata.get("lastName", "")
        email = udata.get("email", "")
        name = f"{first} {last}".strip() or email or user_id
    else:
        name = user_id

    user_cache[user_id] = name
    time.sleep(0.1)
    return name

# -------------------------------------------------
# Prepare Excel
# -------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Spike Incidents"

ws.append([
    "Team Name",
    "Counter ID",
    "Message",
    "Assignee Email",
    "Priority",
    "Status",
    "Source",
    "Created (IST)",
    "ACK At (IST)",
    "Resolved At (IST)",
    "Notes (Only matching NACK date)"
])

# -------------------------------------------------
# Fetch and process incidents
# -------------------------------------------------
for team_name, team_id in teams.items():
    print(f"📡 Fetching incidents for team: {team_name} ({team_id})")

    headers_req = {
        "x-api-key": SPIKE_API_KEY,
        "x-team-id": team_id,
        "Accept": "application/json"
    }

    list_resp = requests.get("https://api.spike.sh/incidents", headers=headers_req)
    if list_resp.status_code != 200:
        print(f"❌ Failed for {team_name} | {list_resp.status_code}")
        continue

    incidents = list_resp.json().get("incidents", list_resp.json())

    for inc in incidents:
        grouped = inc.get("groupedIncident", {})

        nack_dt = utc_to_ist(inc.get("NACK_at"))
        nack_date_str = nack_dt.strftime("%Y-%m-%d") if nack_dt else ""

        assignees = inc.get("assignee", [])
        assignee_emails = ", ".join(
            a.get("email", "") for a in assignees
            if isinstance(a, dict) and a.get("email")
        )

        created_ist = ist_str(utc_to_ist(inc.get("NACK_at") or inc.get("createdAt")))
        ack_ist = ist_str(utc_to_ist(inc.get("ACK_at")))
        res_ist = ist_str(utc_to_ist(inc.get("RES_at")))

        # -------- Notes matching NACK date --------
        filtered_notes = []
        for note in grouped.get("notes", []):
            if not isinstance(note, dict):
                continue
            note_dt = utc_to_ist(note.get("createdAt"))
            if not note_dt:
                continue
            if note_dt.strftime("%Y-%m-%d") != nack_date_str:
                continue
            user_disp = resolve_user(note.get("user"), team_id)
            note_text = note.get("content", "").replace("\n", " ")
            filtered_notes.append(f"{ist_str(note_dt)} | {user_disp}: {note_text}")

        all_notes = "\n".join(filtered_notes)

        ws.append([
            team_name,
            inc.get("counterId"),
            inc.get("message"),
            assignee_emails,
            inc.get("metadata", {}).get("priority"),
            inc.get("status"),
            inc.get("integration", {}).get("name"),
            created_ist,
            ack_ist,
            res_ist,
            all_notes
        ])

# -------------------------------------------------
# Save Excel
# -------------------------------------------------
file_name = f"spike_incidents_report_{datetime.now().date()}.xlsx"
wb.save(file_name)

print(f"\n✅ Report generated: {file_name}")

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
# Convert UTC → IST
# -------------------------------------------------
def utc_to_ist(utc_time_str):
    if not utc_time_str:
        return ""
    try:
        dt = datetime.fromisoformat(utc_time_str.replace("Z", "+00:00"))
        return dt.astimezone(ZoneInfo("Asia/Kolkata")).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return utc_time_str

# -------------------------------------------------
# User cache
# -------------------------------------------------
user_cache = {}

def get_user_info(user_id, team_id):
    """
    Fetch user details for user_id from Spike API and cache result.
    """
    if not user_id:
        return ""

    if user_id in user_cache:
        return user_cache[user_id]

    # Call users endpoint
    resp = requests.get(
        f"https://api.spike.sh/users/{user_id}",
        headers={
            "x-api-key": SPIKE_API_KEY,
            "x-team-id": team_id,
            "Accept": "application/json"
        }
    )
    if resp.status_code == 200:
        u = resp.json()
        # Prefer name, fallback to email, fallback to ID
        name = (u.get("firstName") or "").strip() + " " + (u.get("lastName") or "").strip()
        name = name.strip() or u.get("email") or user_id
    else:
        name = user_id

    user_cache[user_id] = name
    # small sleep to prevent rate limit issues
    time.sleep(0.1)
    return name

# -------------------------------------------------
# Prepare Excel
# -------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Spike Incidents"

headers_row = [
    "Team Name",
    "Counter ID",
    "Message",
    "Assignee Email",
    "Priority",
    "Status",
    "Source",
    "Created Time (IST)",
    "ACK At (IST)",
    "Resolved At (IST)",
    "Notes (User & Time)",
    "Comments (User & Time)"
]
ws.append(headers_row)

# -------------------------------------------------
# Fetch incidents per team
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

    data = list_resp.json()
    incidents_list = data.get("incidents", data)

    for inc in incidents_list:
        grouped = inc.get("groupedIncident", {})

        # Assignee emails
        assignees = inc.get("assignee", [])
        assignee_emails = [
            a.get("email", "")
            for a in assignees
            if isinstance(a, dict) and a.get("email")
        ]
        assignees_str = ", ".join(assignee_emails)

        created_time = inc.get("NACK_at") or inc.get("createdAt")

        # -------- NOTES --------
        notes_output = []
        for note in grouped.get("notes", []):
            if isinstance(note, dict):
                user_field = note.get("user")
                # user_field might be ID string
                user_display = ""
                if isinstance(user_field, str):
                    user_display = get_user_info(user_field, team_id)
                elif isinstance(user_field, dict):
                    user_display = (user_field.get("firstName") or "").strip() + " " + (user_field.get("lastName") or "").strip()
                    user_display = user_display.strip() or user_field.get("email") or ""
                note_time = utc_to_ist(note.get("createdAt"))
                content = note.get("content", "").replace("\n", " ")
                notes_output.append(f"{note_time} | {user_display}: {content}")
            else:
                # if note is simple string
                notes_output.append(str(note))
        all_notes = "\n".join(notes_output)

        # -------- COMMENTS --------
        comments_output = []
        for comm in grouped.get("comments", []):
            if isinstance(comm, dict):
                user_field = comm.get("user")
                if isinstance(user_field, str):
                    user_display = get_user_info(user_field, team_id)
                elif isinstance(user_field, dict):
                    user_display = (user_field.get("firstName") or "").strip() + " " + (user_field.get("lastName") or "").strip()
                    user_display = user_display.strip() or user_field.get("email") or ""
                else:
                    user_display = ""
                comm_time = utc_to_ist(comm.get("createdAt"))
                content = comm.get("content", "").replace("\n", " ")
                comments_output.append(f"{comm_time} | {user_display}: {content}")
            else:
                comments_output.append(str(comm))
        all_comments = "\n".join(comments_output)

        ws.append([
            team_name,
            inc.get("counterId"),
            inc.get("message"),
            assignees_str,
            inc.get("metadata", {}).get("priority"),
            inc.get("status"),
            inc.get("integration", {}).get("name"),
            utc_to_ist(created_time),
            utc_to_ist(inc.get("ACK_at")),
            utc_to_ist(inc.get("RES_at")),
            all_notes,
            all_comments
        ])

# -------------------------------------------------
# Save Excel
# -------------------------------------------------
file_name = f"spike_incidents_report_{datetime.now().date()}.xlsx"
wb.save(file_name)

print(f"\n✅ Report generated: {file_name}")

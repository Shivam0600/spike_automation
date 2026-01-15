import requests
from openpyxl import Workbook
from datetime import datetime
import os
from dotenv import load_dotenv

# Load .env
load_dotenv()

SPIKE_API_KEY = os.getenv("SPIKE_API_KEY")
SPIKE_TEAM_IDS = os.getenv("SPIKE_TEAM_IDS")

if not SPIKE_API_KEY or not SPIKE_TEAM_IDS:
    print("❌ Missing SPIKE_API_KEY or SPIKE_TEAM_IDS in .env")
    exit(1)

TEAM_IDS = [t.strip() for t in SPIKE_TEAM_IDS.split(",")]

url = "https://api.spike.sh/incidents"

wb = Workbook()
ws = wb.active
ws.title = "Spike Incidents"

headers_row = [
    "Team ID",
    "Incident ID",
    "Counter ID",
    "Message",
    "Priority",
    "State",
    "Status",
    "Source",
    "Service",
    "Created Date",
    "ACK At",
    "Resolved At",
    "All Notes",
    "All Comments"
]
ws.append(headers_row)

for team_id in TEAM_IDS:
    print(f"📡 Fetching incidents for team: {team_id}")

    headers = {
        "x-api-key": SPIKE_API_KEY,
        "x-team-id": team_id,
        "Accept": "application/json"
    }

    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        print(f"❌ Failed for team {team_id} | {response.status_code}")
        continue

    data = response.json()

    if isinstance(data, dict) and "incidents" in data:
        incidents_list = data["incidents"]
    elif isinstance(data, list):
        incidents_list = data
    else:
        print(f"❌ Unexpected response for team {team_id}")
        continue

    for inc in incidents_list:
        # -------- NOTES --------
        notes_list = []
        grouped = inc.get("groupedIncident", {})

        for note in grouped.get("notes", []):
            if isinstance(note, dict):
                user = note.get("user", "")
                content = note.get("content", "").replace("\n", " ")
                created_at = note.get("createdAt", "")
                notes_list.append(f"{created_at} | {user}: {content}")

        all_notes = "\n".join(notes_list)

        # -------- COMMENTS --------
        comments_list = []
        for comm in grouped.get("comments", []):
            if isinstance(comm, dict):
                user = comm.get("user", "")
                content = comm.get("content", "").replace("\n", " ")
                created_at = comm.get("createdAt", "")
                comments_list.append(f"{created_at} | {user}: {content}")

        all_comments = "\n".join(comments_list)

        ws.append([
            team_id,
            inc.get("_id"),
            inc.get("counterId"),
            inc.get("message"),
            inc.get("metadata", {}).get("priority"),
            inc.get("metadata", {}).get("state"),
            inc.get("status"),
            inc.get("integration", {}).get("name"),
            ",".join(inc.get("metadata", {}).get("impactedEntities", [])),
            inc.get("createdAt"),
            inc.get("ACK_at"),
            inc.get("RES_at"),
            all_notes,
            all_comments
        ])

file_name = f"spike_incidents_report_ALL_TEAMS_{datetime.now().date()}.xlsx"
wb.save(file_name)

print(f"\n✅ Incident report for ALL teams generated: {file_name}")

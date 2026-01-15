import os
import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from datetime import datetime, timezone, timedelta

# =====================================================
# ENV
# =====================================================
load_dotenv()
API_KEY = os.getenv("SPIKE_API_KEY")

TEAMS = {
    "Spinomenal": os.getenv("TEAM_SPINOMENAL"),
    "Mobideo": os.getenv("TEAM_MOBIDEO"),
    "Duve": os.getenv("TEAM_DUVE"),
    "Delek US": os.getenv("TEAM_DELEK_US"),
    "Qbiq": os.getenv("TEAM_QBIQ"),
    "The5ers": os.getenv("TEAM_THE5ERS"),
    "Filuet": os.getenv("TEAM_FILUET"),
}

URL = "https://api.spike.sh/incidents"

# =====================================================
# DATE FILTER (RECEIVED DATE)
# =====================================================
FILTER_DATE = "2026-01-13"   # YYYY-MM-DD
FILTER_DATE_OBJ = datetime.strptime(FILTER_DATE, "%Y-%m-%d").date()

IST = timezone(timedelta(hours=5, minutes=30))

# =====================================================
# EXCEL SETUP
# =====================================================
wb = Workbook()
wb.remove(wb.active)

# =====================================================
# FETCH & EXPORT
# =====================================================
for team_name, team_id in TEAMS.items():
    print(f"📡 Fetching incidents for {team_name}")

    response = requests.get(
        URL,
        headers={
            "x-api-key": API_KEY,
            "x-team-id": team_id,
            "Accept": "*/*"
        }
    )
    response.raise_for_status()
    data = response.json()

    incidents = (
        data.get("result", {}).get("incidents")
        or data.get("incidents")
        or []
    )

    print(f"   → Total incidents: {len(incidents)}")

    filtered = []
    for inc in incidents:
        created_at = inc.get("createdAt")
        if not created_at:
            continue

        created_dt = datetime.fromisoformat(
            created_at.replace("Z", "+00:00")
        ).astimezone(IST)

        if created_dt.date() == FILTER_DATE_OBJ:
            filtered.append(inc)

    print(f"   → Incidents on {FILTER_DATE}: {len(filtered)}")

    if not filtered:
        continue

    ws = wb.create_sheet(team_name[:31])
    ws.append([
        "Incident ID",
        "Title",
        "Notes",
        "Received Time",
        "Acknowledged Time",
        "Resolved Time"
    ])

    for inc in filtered:
        notes_list = (
            inc.get("groupedIncident", {})
               .get("notes", [])
        )
        notes_text = "; ".join(n.get("content", "") for n in notes_list)

        ws.append([
            inc.get("counterId"),   # ✅ Human-readable ID
            inc.get("message"),
            notes_text,
            inc.get("createdAt"),
            inc.get("ACK_at"),
            inc.get("RES_at"),
        ])

# =====================================================
# SAVE
# =====================================================
filename = f"Spike_Incidents_{FILTER_DATE}.xlsx"
wb.save(filename)
print(f"\n✅ Excel generated successfully: {filename}")

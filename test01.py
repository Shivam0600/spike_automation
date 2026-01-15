import requests
from openpyxl import Workbook
from datetime import datetime
import os
from dotenv import load_dotenv

# load variables from .env
load_dotenv()

SPIKE_API_KEY = os.getenv("SPIKE_API_KEY")
SPIKE_TEAM_ID = os.getenv("SPIKE_TEAM_ID")

if not SPIKE_API_KEY or not SPIKE_TEAM_ID:
    print("❌ Missing SPIKE_API_KEY or SPIKE_TEAM_ID in .env")
    exit(1)

url = "https://api.spike.sh/incidents"

headers = {
    "x-api-key": SPIKE_API_KEY,
    "x-team-id": SPIKE_TEAM_ID,
    "Accept": "application/json"
}

response = requests.get(url, headers=headers)

# Handle invalid API or unauthorized
if response.status_code != 200:
    print("❌ Failed to fetch incidents")
    print("Status code:", response.status_code)
    print("Response:", response.text)
    exit(1)

data = response.json()

# Detect actual incidents list
if isinstance(data, dict) and "incidents" in data:
    incidents_list = data["incidents"]
elif isinstance(data, list):
    incidents_list = data
else:
    print("❌ Unexpected API response structure")
    print(data)
    exit(1)

wb = Workbook()
ws = wb.active
ws.title = "Spike Incidents"

headers_row = [
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
    "Resolved At"
]
ws.append(headers_row)

for inc in incidents_list:
    ws.append([
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
        inc.get("RES_at")
    ])

file_name = f"spike_incidents_{datetime.now().date()}.xlsx"
wb.save(file_name)

print(f"✅ Incident report generated: {file_name}")

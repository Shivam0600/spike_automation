import os
import requests
from dotenv import load_dotenv
from openpyxl import Workbook

load_dotenv()
API_KEY = os.getenv("SPIKE_API_KEY")

url = "https://api.spike.sh/teams/get-all-teams"

headers = {
    "x-api-key": API_KEY,
    "Accept": "application/json"
}

response = requests.get(url, headers=headers)
response.raise_for_status()

json_data = response.json()

# 🔍 DEBUG (remove later)
print("API Response Keys:", json_data.keys())

# ✅ FIX: find correct teams list
teams = (
    json_data.get("data")
    or json_data.get("teams")
    or json_data.get("result")
    or []
)

if not teams:
    print("⚠️ No teams found in API response")
    exit(0)

# ===============================
# Excel
# ===============================
wb = Workbook()
ws = wb.active
ws.title = "Teams"

ws.append(["Team Name", "Team ID"])

for team in teams:
    ws.append([
        team.get("name"),
        team.get("_id")
    ])

wb.save("spike_teams.xlsx")

print("✅ Teams exported successfully")

# AttendSmart IoT 2025 V.2.1
# By Senuka Damketh
# Sri Bodhiraja College Embilipitiya

from flask import Flask, request, jsonify
from colorama import Fore
from datetime import datetime
import gspread
import json
import os
import requests
from oauth2client.service_account import ServiceAccountCredentials

# Colorful Banner
print()
print(Fore.RED +     "      █████╗ ████████╗████████╗███████╗███╗   ██╗██████╗     ███████╗███╗   ███╗ █████╗ ██████╗ ████████╗    ")
print(Fore.YELLOW +  "     ██╔══██╗╚══██╔══╝╚══██╔══╝██╔════╝████╗  ██║██╔══██╗    ██╔════╝████╗ ████║██╔══██╗██╔══██╗╚══██╔══╝     ")
print(Fore.GREEN +   "     ███████║   ██║      ██║   █████╗  ██╔██╗ ██║██║  ██║    ███████╗██╔████╔██║███████║██████╔╝   ██║        ")
print(Fore.BLUE +    "     ██╔══██║   ██║      ██║   ██╔══╝  ██║╚██╗██║██║  ██║    ╚════██║██║╚██╔╝██║██╔══██║██╔══██╗   ██║        ")
print(Fore.CYAN +    "     ██║  ██║   ██║      ██║   ███████╗██║ ╚████║██████╔╝    ███████║██║ ╚═╝ ██║██║  ██║██║  ██║   ██║        ")
print(Fore.MAGENTA + "     ╚═╝  ╚═╝   ╚═╝      ╚═╝   ╚══════╝╚═╝  ╚═══╝╚═════╝     ╚══════╝╚═╝     ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝   ╚═╝        ")
print()
print(Fore.CYAN + "----------------------------------------")
print(Fore.CYAN + "AttendSmart IoT 2025 V.2.1")
print(Fore.CYAN + "By Senuka Damketh")
print(Fore.CYAN + "Sri Bodhiraja College Embilipitiya")
print(Fore.CYAN + "----------------------------------------\n\n")

# Load or ask for credentials path
cred_path_file = "credentials_path.txt"

if os.path.exists(cred_path_file):
    with open(cred_path_file, "r") as f:
        input_path = f.read().strip()
else:
    print(Fore.BLUE + "Enter Google Sheets API Credentials JSON file path: ", end="")
    input_path = input().strip()
    with open(cred_path_file, "w") as f:
        f.write(input_path)

# Google Sheets setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(input_path, scope)
client = gspread.authorize(creds)

# Flask app
app = Flask(__name__)

# Load or initialize seen_cards
seen_cards_file = "seen_cards.json"
if os.path.exists(seen_cards_file):
    with open(seen_cards_file, "r") as f:
        seen_cards = json.load(f)
else:
    seen_cards = {}

def save_seen_cards():
    with open(seen_cards_file, "w") as f:
        json.dump(seen_cards, f)

@app.route("/post_data", methods=["POST"])
def post_data():
    data = request.get_json()
    uid = data.get("uid", "").upper()
    device_id = data.get("device_id", "")

    if not uid or not device_id:
        return jsonify({"status": "error", "message": "Missing UID or device_id"}), 400

    now = datetime.now()
    date_str = str(now.day)
    time_str = now.strftime("%H:%M")
    month_year = now.strftime("%B %Y")
    today_key = now.strftime("%Y-%m-%d")
    hour = now.hour

    print(f"Received UID {uid} from {device_id} at {time_str}")

    # Load student register
    try:
        register_sheet = client.open("student register").sheet1
        students = register_sheet.get_all_records()
    except Exception as e:
        return jsonify({"status": "error", "message": f"Failed to open student register: {e}"}), 500

    student = next((s for s in students if s["UID"].upper() == uid), None)

    if not student:
        return jsonify({"status": "not_found", "message": f"UID {uid} not found in register"}), 404

    name = student["Name"]
    ntfy_url = student["ntfy URL"]
    sheet_name = f"{month_year}-{device_id}"

    # Open or create monthly spreadsheet
    try:
        try:
            month_sheet = client.open(sheet_name)
        except gspread.exceptions.SpreadsheetNotFound:
            month_sheet = client.create(sheet_name)
            print(f"Created new spreadsheet: {sheet_name}")
    except Exception as e:
        return jsonify({"status": "error", "message": f"Error accessing/creating spreadsheet: {e}"}), 500

    # Open or create worksheet for today's date
    try:
        try:
            sheet = month_sheet.worksheet(date_str)
        except gspread.exceptions.WorksheetNotFound:
            sheet = month_sheet.add_worksheet(title=date_str, rows="100", cols="4")
            sheet.append_row(["Name", "UID", "Arrival Time", "Leave Time"])
    except Exception as e:
        return jsonify({"status": "error", "message": f"Error accessing/creating worksheet: {e}"}), 500

    # Get current records to check for existing UID
    try:
        records = sheet.get_all_values()
    except Exception as e:
        return jsonify({"status": "error", "message": f"Failed to read sheet: {e}"}), 500

    # Check if UID is already in sheet
    row_index = None
    for i, row in enumerate(records):
        if len(row) > 1 and row[1].upper() == uid:
            row_index = i + 1  # gspread uses 1-based index
            break

    # Update or create entry
    try:
        if row_index:
            current_row = records[row_index - 1]
            if hour < 12 and (len(current_row) < 3 or not current_row[2]):
                sheet.update_cell(row_index, 3, time_str)  # Arrival
            elif hour >= 12 and (len(current_row) < 4 or not current_row[3]):
                sheet.update_cell(row_index, 4, time_str)  # Leave
            else:
                print(f"UID {uid} already logged for this time.")
                return jsonify({"status": "duplicate"}), 200
        else:
            # Append new row with correct column filled
            arrival = time_str if hour < 12 else ""
            leave = time_str if hour >= 12 else ""
            sheet.append_row([name, uid, arrival, leave])
    except Exception as e:
        return jsonify({"status": "error", "message": f"Failed to update/append row: {e}"}), 500

    # Send notification
    status = "arrived to" if hour < 12 else "left"
    message = f"{name} has {status} Bodhiraja college at {time_str}"

    try:
        requests.post(ntfy_url, data=message.encode("utf-8"))
        print("Notification sent.")
    except Exception as e:
        print("Notification failed:", e)

    # Mark UID as seen today
    if today_key not in seen_cards:
        seen_cards[today_key] = {}
    seen_cards[today_key][uid] = status
    save_seen_cards()

    return jsonify({"status": "success", "name": name})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

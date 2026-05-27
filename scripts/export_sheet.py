"""
export_sheet.py
Exports every sheet from the Stone Arch Golf League Google Spreadsheet into:
  1. data/Stone_Arch_Golf_League_Data_2026.xlsx  (full workbook)
  2. data.json  (all sheets as JSON, served via GitHub Pages for Claude to read)
"""

import os
import json
import base64
from google.oauth2 import service_account
from googleapiclient.discovery import build
from openpyxl import Workbook

SHEET_ID = os.environ["SHEET_ID"]
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

def get_credentials():
    raw = os.environ["GOOGLE_CREDENTIALS"]
    try:
        info = json.loads(raw)
    except json.JSONDecodeError:
        info = json.loads(base64.b64decode(raw).decode("utf-8"))
    return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)

def main():
    creds = get_credentials()
    service = build("sheets", "v4", credentials=creds)
    spreadsheet = service.spreadsheets()

    # Get list of all sheet names
    meta = spreadsheet.get(spreadsheetId=SHEET_ID).execute()
    sheets = meta.get("sheets", [])

    wb = Workbook()
    wb.remove(wb.active)
    all_data = {}

    for sheet in sheets:
        title = sheet["properties"]["title"]
        print(f"  Exporting: {title}")

        result = spreadsheet.values().get(
            spreadsheetId=SHEET_ID,
            range=title
        ).execute()

        values = result.get("values", [])

        # Add to xlsx
        ws = wb.create_sheet(title=title)
        for row in values:
            ws.append(row)

        # Add to JSON dict
        all_data[title] = values

    # Save xlsx
    os.makedirs("data", exist_ok=True)
    xlsx_path = "data/Stone_Arch_Golf_League_Data_2026.xlsx"
    wb.save(xlsx_path)
    print(f"Saved xlsx: {xlsx_path}")

    # Save JSON to repo root (served by GitHub Pages)
    json_path = "data.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False)
    print(f"Saved JSON: {json_path} ({len(all_data)} sheets)")

if __name__ == "__main__":
    main()

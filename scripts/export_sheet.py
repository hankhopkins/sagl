"""
export_sheet.py
Exports every sheet from the Stone Arch Golf League Google Spreadsheet
into a single .xlsx file saved at data/Stone_Arch_Golf_League_Data_2026.xlsx
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
    # Support both raw JSON and base64-encoded JSON
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
    wb.remove(wb.active)  # remove default empty sheet

    for sheet in sheets:
        title = sheet["properties"]["title"]
        print(f"  Exporting: {title}")

        result = spreadsheet.values().get(
            spreadsheetId=SHEET_ID,
            range=title
        ).execute()

        values = result.get("values", [])
        ws = wb.create_sheet(title=title)

        for row in values:
            ws.append(row)

    os.makedirs("data", exist_ok=True)
    output_path = "data/Stone_Arch_Golf_League_Data_2026.xlsx"
    wb.save(output_path)
    print(f"\nSaved to {output_path} ({len(sheets)} sheets)")

if __name__ == "__main__":
    main()

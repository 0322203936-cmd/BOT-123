import sys
import json
from urllib.parse import quote
from sharepoint_sync import (
    graph_token,
    resolve_sharepoint_item_by_url,
    GRAPH_URL,
    graph_headers
)
from excel_range_sync import graph_request

PLAN_COSECHA_URL = "https://pacificafarms.sharepoint.com/:x:/r/sites/DocCampos/_layouts/15/Doc.aspx?sourcedoc=%7BB574F211-4861-4031-8C8E-03448B593DA2%7D&file=Plan%20de%20cosecha%202025.xlsx&action=default&mobileredirect=true"

def main():
    token = graph_token()
    item_plan = resolve_sharepoint_item_by_url(token, PLAN_COSECHA_URL)
    plan_drive_id = item_plan["parentReference"]["driveId"]
    plan_workbook_url = f"{GRAPH_URL}/drives/{plan_drive_id}/items/{item_plan['id']}/workbook"
    headers = {**graph_headers(token), "Content-Type": "application/json"}
    
    plan_session_res = graph_request(
        "POST",
        f"{plan_workbook_url}/createSession",
        headers,
        json={"persistChanges": False},
        timeout=60,
    ).json()
    plan_session_headers = {**headers, "workbook-session-id": plan_session_res["id"]}
    
    try:
        worksheets_data = graph_request("GET", f"{plan_workbook_url}/worksheets", plan_session_headers).json()
        sheetnames = [ws["name"] for ws in worksheets_data.get("value", [])]
        print("Sheets:", sheetnames)
        
        cosecha_sheets = [s for s in sheetnames if s.startswith("P Cosecha ")]
        latest_sheet_name = sorted(cosecha_sheets)[-1]
        print("Latest:", latest_sheet_name)
        
        quoted_sheet = quote(latest_sheet_name)
        range_data = graph_request("GET", f"{plan_workbook_url}/worksheets('{quoted_sheet}')/range(address='A1:X10')", plan_session_headers).json()
        values = range_data.get("values", [])
        
        print("Row 4:", values[3])
    finally:
        try:
            graph_request("POST", f"{plan_workbook_url}/closeSession", plan_session_headers)
        except:
            pass

if __name__ == '__main__':
    main()

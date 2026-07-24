import sys
import re

from sharepoint_sync import (
    graph_token,
    resolve_sharepoint_item_by_url,
    GRAPH_URL,
    graph_headers
)
from excel_range_sync import graph_request

REQ_PROY_URL = "https://pacificafarms.sharepoint.com/:x:/r/sites/requerimientovsproyeccion/_layouts/15/Doc.aspx?sourcedoc=%7B277A76AA-508A-47F8-8A4A-F19D46660D65%7D&file=Requerimiento%20vs%20proyeccion%20Test.xlsm&action=default&mobileredirect=true"

def col_letter_to_index(letter):
    idx = 0
    for char in letter.upper():
        idx = idx * 26 + (ord(char) - ord('A') + 1)
    return idx - 1

def main():
    print("Obteniendo datos de SharePoint...")
    token = graph_token()
    item_req = resolve_sharepoint_item_by_url(token, REQ_PROY_URL)
    drive_id = item_req["parentReference"]["driveId"]
    workbook_url = f"{GRAPH_URL}/drives/{drive_id}/items/{item_req['id']}/workbook"
    headers = {**graph_headers(token), "Content-Type": "application/json"}
    
    session = graph_request("POST", f"{workbook_url}/createSession", headers, json={"persistChanges": False}, timeout=60).json()
    session_headers = {**headers, "workbook-session-id": session["id"]}
    
    try:
        used_range_res = graph_request("GET", f"{workbook_url}/worksheets/DataProy/usedRange", session_headers, timeout=120).json()
        address = used_range_res.get("address", "")
        values = used_range_res.get("values", [])
        
        match = re.search(r'!([A-Za-z]+)(\d+)', address)
        start_col_str = match.group(1).upper() if match else "A"
        start_row = int(match.group(2)) if match else 1
        
        start_col_idx = col_letter_to_index(start_col_str)
        col_h = col_letter_to_index("H") - start_col_idx
        
        # Simulate logic
        in_corte_block = False
        seen_compra = False
        corte_count = 0
        num_flowers = 30 # example
        total_corte_rows_needed = num_flowers * 8
        
        target_rows = []
        rows_to_clear = []
        
        print(f"Address: {address}, start_row: {start_row}, total_values: {len(values)}")
        
        for idx, row_data in enumerate(values):
            if len(row_data) > col_h and col_h >= 0:
                desc = str(row_data[col_h]).strip().upper()
            else:
                desc = ""
                
            if desc == "COMPRA":
                seen_compra = True
                in_corte_block = False
            elif seen_compra and desc != "COMPRA" and corte_count < total_corte_rows_needed:
                in_corte_block = True
                excel_row = start_row + idx
                target_rows.append(excel_row)
                corte_count += 1
            elif seen_compra and desc == "CORTE" and corte_count >= total_corte_rows_needed:
                excel_row = start_row + idx
                rows_to_clear.append(excel_row)
            else:
                in_corte_block = False
                
        print(f"Target rows for CORTE (Count: {len(target_rows)}): {target_rows[:5]} ... {target_rows[-5:]}")
        print(f"Leftover CORTE rows to clear (Count: {len(rows_to_clear)}): {rows_to_clear[:5]} ... {rows_to_clear[-5:]}")
        
    finally:
        try:
            graph_request("POST", f"{workbook_url}/closeSession", session_headers, timeout=30)
        except Exception:
            pass

if __name__ == '__main__':
    main()

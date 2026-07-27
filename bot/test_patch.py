import sys
from sharepoint_sync import (
    graph_token,
    resolve_sharepoint_item_by_url,
    GRAPH_URL,
    graph_headers
)
from excel_range_sync import graph_request

REQ_PROY_URL = "https://pacificafarms.sharepoint.com/:x:/r/sites/requerimientovsproyeccion/_layouts/15/Doc.aspx?sourcedoc=%7BB6111299-1373-4717-A2B7-3D507AD77A8A%7D&file=Requerimiento%20vs%20proyeccion.xlsm&action=default&mobileredirect=true"

def main():
    token = graph_token()
    item_req = resolve_sharepoint_item_by_url(token, REQ_PROY_URL)
    drive_id = item_req["parentReference"]["driveId"]
    workbook_url = f"{GRAPH_URL}/drives/{drive_id}/items/{item_req['id']}/workbook"
    headers = {**graph_headers(token), "Content-Type": "application/json"}
    
    session = graph_request("POST", f"{workbook_url}/createSession", headers, json={"persistChanges": True}, timeout=60).json()
    session_headers = {**headers, "workbook-session-id": session["id"]}
    
    try:
        # First read the cell S233
        res = graph_request("GET", f"{workbook_url}/worksheets/DataProy/range(address='S233')", session_headers).json()
        print(f"Old value in S233: {res.get('values')}")
        
        # Now try to patch it to 999
        print("Patching S233 to 999...")
        patch_res = graph_request("PATCH", f"{workbook_url}/worksheets/DataProy/range(address='S233')", session_headers, json={"values": [[999]]}).json()
        print(f"Patch response values: {patch_res.get('values')}")
        
        # Read it again
        res2 = graph_request("GET", f"{workbook_url}/worksheets/DataProy/range(address='S233')", session_headers).json()
        print(f"New value in S233: {res2.get('values')}")
        
    finally:
        try:
            graph_request("POST", f"{workbook_url}/closeSession", session_headers, timeout=30)
        except Exception:
            pass

if __name__ == '__main__':
    main()

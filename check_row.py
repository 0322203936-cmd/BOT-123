import sys
sys.path.append('bot')
from sharepoint_sync import graph_token, GRAPH_URL, resolve_sharepoint_item_by_url, graph_headers
from excel_range_sync import graph_request

token = graph_token()
item = resolve_sharepoint_item_by_url(token, 'https://pacificafarms.sharepoint.com/:x:/r/sites/requerimientovsproyeccion/_layouts/15/Doc.aspx?sourcedoc=%7B277A76AA-508A-47F8-8A4A-F19D46660D65%7D&file=Requerimiento%20vs%20proyeccion%20Test.xlsm&action=default&mobileredirect=true')
url = f"{GRAPH_URL}/drives/{item['parentReference']['driveId']}/items/{item['id']}/workbook"
sess = graph_request('POST', f'{url}/createSession', {**graph_headers(token), 'Content-Type': 'application/json'}, json={'persistChanges': True}).json()
sh = {**graph_headers(token), 'workbook-session-id': sess['id']}
res = graph_request('GET', f"{url}/worksheets/DataProy/range(address='A1893:Z1893')", sh).json()
print("Values:", res.get('values'))
print("Formulas:", res.get('formulas'))

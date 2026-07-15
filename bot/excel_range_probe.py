import requests

from sharepoint_sync import GRAPH_URL, graph_headers, graph_token, resolve_sharepoint_item


def request(response: requests.Response) -> requests.Response:
    if not response.ok:
        try:
            detail = response.json().get("error", {}).get("message", "")
        except ValueError:
            detail = response.text
        raise RuntimeError(
            f"Microsoft Graph respondio {response.status_code}: {detail}"
        )
    return response


def run() -> None:
    token = graph_token()
    item = resolve_sharepoint_item(token)
    drive_id = item["parentReference"]["driveId"]
    workbook_url = (
        f"{GRAPH_URL}/drives/{drive_id}/items/{item['id']}/workbook"
    )
    headers = {**graph_headers(token), "Content-Type": "application/json"}

    session_response = request(
        requests.post(
            f"{workbook_url}/createSession",
            headers=headers,
            json={"persistChanges": False},
            timeout=60,
        )
    )
    session_id = session_response.json()["id"]
    session_headers = {**headers, "workbook-session-id": session_id}
    try:
        used_range = request(
            requests.get(
                f"{workbook_url}/worksheets/PegarData/usedRange",
                headers=session_headers,
                timeout=60,
            )
        ).json()
        table_range = request(
            requests.get(
                f"{workbook_url}/tables/Table2/range",
                headers=session_headers,
                timeout=60,
            )
        ).json()
        print(
            "EXCEL_RANGE_SESSION_OK "
            f"used={used_range.get('address')} "
            f"table={table_range.get('address')} "
            f"rows={table_range.get('rowCount')}"
        )
    finally:
        requests.post(
            f"{workbook_url}/closeSession",
            headers=session_headers,
            timeout=30,
        )


if __name__ == "__main__":
    run()

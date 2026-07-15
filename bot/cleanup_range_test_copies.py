import time

import requests

from sharepoint_sync import GRAPH_URL, graph_headers, graph_token, resolve_sharepoint_item


PREFIX = "Reunion BOT RANGO TEST "


def run() -> None:
    token = graph_token()
    main_item = resolve_sharepoint_item(token)
    drive_id = main_item["parentReference"]["driveId"]
    parent_id = main_item["parentReference"]["id"]
    response = requests.get(
        f"{GRAPH_URL}/drives/{drive_id}/items/{parent_id}/children",
        headers=graph_headers(token),
        params={"$select": "id,name"},
        timeout=60,
    )
    response.raise_for_status()
    items = [item for item in response.json().get("value", []) if item["name"].startswith(PREFIX)]
    for item in items:
        for attempt in range(1, 7):
            delete = requests.delete(
                f"{GRAPH_URL}/drives/{drive_id}/items/{item['id']}",
                headers=graph_headers(token),
                timeout=60,
            )
            if delete.status_code == 423 and attempt < 6:
                time.sleep(10)
                continue
            delete.raise_for_status()
            print(f"COPIA_TEMPORAL_ELIMINADA nombre={item['name']}")
            break
    print(f"LIMPIEZA_OK archivos={len(items)}")


if __name__ == "__main__":
    run()

import os
from pathlib import Path

import requests

from sharepoint_sync import GRAPH_URL, graph_headers, graph_token, resolve_sharepoint_item


def run() -> None:
    token = graph_token()
    item = resolve_sharepoint_item(token)
    drive_id = item["parentReference"]["driveId"]
    item_id = item["id"]
    response = requests.get(
        f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/versions?$top=30",
        headers=graph_headers(token),
        timeout=30,
    )
    response.raise_for_status()
    for version in response.json().get("value", []):
        modified_by = version.get("lastModifiedBy", {})
        actor = modified_by.get("user", {}).get("displayName") or modified_by.get(
            "application", {}
        ).get("displayName", "desconocido")
        print(
            "VERSION "
            f"id={version.get('id')} "
            f"fecha={version.get('lastModifiedDateTime')} "
            f"tamano={version.get('size')} "
            f"actual={version.get('lastVersion', False)} "
            f"autor={actor}"
        )

    version_id = os.environ.get("SHAREPOINT_VERSION_ID", "").strip()
    if version_id:
        content = requests.get(
            f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/versions/{version_id}/content",
            headers=graph_headers(token),
            timeout=180,
        )
        content.raise_for_status()
        destination = Path("artifacts/recovery") / f"Reunion_version_{version_id}.xlsm"
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(content.content)
        print(f"VERSION_DESCARGADA id={version_id} ruta={destination}")

    restore_version_id = os.environ.get("SHAREPOINT_RESTORE_VERSION_ID", "").strip()
    if restore_version_id:
        restored = requests.post(
            f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/versions/"
            f"{restore_version_id}/restoreVersion",
            headers=graph_headers(token),
            timeout=60,
        )
        restored.raise_for_status()
        current = requests.get(
            f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/content",
            headers=graph_headers(token),
            timeout=180,
        )
        current.raise_for_status()
        destination = Path("artifacts/recovery") / "Reunion_restaurado.xlsm"
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(current.content)
        print(f"VERSION_RESTAURADA id={restore_version_id} ruta={destination}")


if __name__ == "__main__":
    run()

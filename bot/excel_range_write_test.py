import os
from pathlib import Path
from urllib.parse import quote

import requests

from excel_range_sync import update_pegar_data_range
from sharepoint_sync import (
    GRAPH_URL,
    download_sharepoint_workbook,
    graph_headers,
    graph_token,
    resolve_sharepoint_item,
)


def run() -> None:
    reports = list(Path("artifacts/source").rglob("detalleFlorFiltro_formateado.xlsx"))
    if len(reports) != 1:
        raise RuntimeError(f"Se esperaba un reporte formateado y se encontraron {len(reports)}.")

    token = graph_token()
    main_item = resolve_sharepoint_item(token)
    main_path = download_sharepoint_workbook(token, main_item)
    drive_id = main_item["parentReference"]["driveId"]
    parent_id = main_item["parentReference"]["id"]
    suffix = os.environ.get("GITHUB_RUN_ID", "local")
    filename = f"Reunion BOT RANGO TEST {suffix}.xlsm"
    upload_url = (
        f"{GRAPH_URL}/drives/{drive_id}/items/{parent_id}:/"
        f"{quote(filename, safe='')}:/content"
    )
    response = requests.put(
        upload_url,
        headers={**graph_headers(token), "Content-Type": "application/octet-stream"},
        data=main_path.read_bytes(),
        timeout=180,
    )
    response.raise_for_status()
    test_item = response.json()
    try:
        count = update_pegar_data_range(token, test_item, reports[0])
        print(f"EXCEL_RANGE_ISOLATED_TEST_OK filas={count} archivo={filename}", flush=True)
    finally:
        delete = requests.delete(
            f"{GRAPH_URL}/drives/{drive_id}/items/{test_item['id']}",
            headers=graph_headers(token),
            timeout=60,
        )
        delete.raise_for_status()
        print("Copia temporal eliminada.", flush=True)


if __name__ == "__main__":
    run()

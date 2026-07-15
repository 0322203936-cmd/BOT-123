from pathlib import Path
from urllib.parse import quote

import requests

from sharepoint_sync import GRAPH_URL, graph_headers, graph_token, resolve_sharepoint_item


def run() -> None:
    token = graph_token()
    main_item = resolve_sharepoint_item(token)
    drive_id = main_item["parentReference"]["driveId"]
    parent_id = main_item["parentReference"]["id"]
    filename = "Reunion 1-2-3 Test BOT PRUEBA.xlsm"
    encoded_name = quote(filename, safe="")
    item_url = f"{GRAPH_URL}/drives/{drive_id}/items/{parent_id}:/{encoded_name}"

    metadata = requests.get(item_url, headers=graph_headers(token), timeout=30)
    metadata.raise_for_status()
    item = metadata.json()

    output_dir = Path("artifacts/test-copy-validation")
    output_dir.mkdir(parents=True, exist_ok=True)
    workbook = requests.get(
        f"{GRAPH_URL}/drives/{drive_id}/items/{item['id']}/content",
        headers=graph_headers(token),
        timeout=180,
    )
    workbook.raise_for_status()
    workbook_path = output_dir / filename
    workbook_path.write_bytes(workbook.content)

    pdf = requests.get(
        f"{GRAPH_URL}/drives/{drive_id}/items/{item['id']}/content?format=pdf",
        headers=graph_headers(token),
        timeout=300,
    )
    pdf.raise_for_status()
    pdf_path = output_dir / "Reunion BOT PRUEBA.pdf"
    pdf_path.write_bytes(pdf.content)
    print(
        f"OFFICE_CONVERSION_OK xlsm={workbook_path.stat().st_size} "
        f"pdf={pdf_path.stat().st_size} item={item['id']}"
    )


if __name__ == "__main__":
    run()

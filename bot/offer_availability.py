from __future__ import annotations

from pathlib import Path

from playwright.sync_api import sync_playwright

from inventory_box_transform import create_single_sheet_workbook
from inventory_boxes import (
    capture,
    delete_all_inventory,
    login_kometsales,
    open_boxes,
    required_secret,
    upload_boxes,
)
from sharepoint_sync import (
    download_sharepoint_file,
    graph_token,
    resolve_sharepoint_item_by_url,
)


ARTIFACTS_DIR = Path("artifacts/offer_availability")
SOURCE_FILENAME = "availability-source.xlsx"
UPLOAD_FILENAME = "availability-only.xlsx"


def prepare_availability_workbook() -> Path:
    sharepoint_url = required_secret("SHAREPOINT_OFFER_URL")
    token = graph_token()
    item = resolve_sharepoint_item_by_url(token, sharepoint_url)
    source = download_sharepoint_file(token, item, SOURCE_FILENAME)
    destination = ARTIFACTS_DIR / UPLOAD_FILENAME
    create_single_sheet_workbook(
        source,
        destination,
        sheet_name="Availability",
        normalize_available_from_dates=True,
    )
    print(f"Archivo preparado con una sola hoja: {destination}", flush=True)
    return destination


def run() -> None:
    komet_user = required_secret("KOMET_USER")
    komet_password = required_secret("KOMET_PASSWORD")
    workbook_path = prepare_availability_workbook()

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            accept_downloads=True,
        )
        page = context.new_page()
        try:
            login_kometsales(page, komet_user, komet_password)
            capture(page, "ofrecer_00_sesion_iniciada.png")
            open_boxes(page)
            capture(page, "ofrecer_01_inventario_actual.png")

            # delete_all_inventory ya contempla el caso en que Komet está vacío:
            # no intenta borrar y permite continuar directamente con la carga.
            delete_all_inventory(page)
            upload_boxes(page, workbook_path)
            capture(page, "ofrecer_05_finalizado.png")
            print(
                "Availability se cargó en Komet sin ejecutar reglas de inventario, "
                "fechas o correo.",
                flush=True,
            )
        except Exception:
            capture(page, "ofrecer_99_error.png")
            raise
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    run()

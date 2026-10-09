"""Independent, review-only artifact pass after the Facturas Komet job."""

from __future__ import annotations

import json
import os
import re
from datetime import date
from pathlib import Path

from playwright.sync_api import sync_playwright

if __package__:
    from . import facturas_komet as komet
    from . import posco_order_import as posco_import
    from . import posco_order_transform as transform
else:
    import facturas_komet as komet
    import posco_order_import as posco_import
    import posco_order_transform as transform


ARTIFACTS_DIR = Path("artifacts/facturas_komet/posco_orders")
REPORT_PATH = ARTIFACTS_DIR / "report.json"
LEDGER_PATH = Path(os.environ.get("KOMET_POSCO_TRANSFORMED_PATH", "bot/data/posco_transformados.json"))
STAGE_LEDGER_PATH = Path(os.environ.get("KOMET_POSCO_STAGE_PATH", "bot/data/posco_cargas.json"))
ORDERS_URL = "https://app.kometsales.com/orderSummary/list.do#st"


def verify_order_number(selected: str, downloaded: str) -> None:
    selected_number = re.search(r"\d{6}", selected)
    downloaded_number = re.search(r"\d{6}", downloaded) or re.fullmatch(r"\d{1,6}", downloaded)
    if not selected_number or not downloaded_number or selected_number.group() != downloaded_number.group().zfill(6):
        raise transform.TransformationError(f"El XLS descargado es de la orden {downloaded}, no de {selected}.")


def collect_filtered_orders(page) -> list[dict[str, str]]:
    seen_pages: set[tuple[str, ...]] = set()
    found: list[dict[str, str]] = []
    for _page_number in range(1, 101):
        orders = komet.collect_order_rows(page)
        signature = komet.page_signature(orders)
        if not orders or signature in seen_pages:
            break
        seen_pages.add(signature)
        found.extend(orders)
        if not komet.click_next_page(page):
            break
    return found


def restore_filtered_orders(page, from_date: date, until_date: date) -> None:
    # An unsuccessful export can navigate away from the jqGrid. Rebuild the view
    # instead of trying stale row selectors from the previous page.
    page.goto(ORDERS_URL, wait_until="domcontentloaded", timeout=30_000)
    if "/orderSummary/list.do" not in page.url:
        raise RuntimeError(f"Komet no volvió al listado de órdenes: {page.url}")
    page.locator("#txtFromDateTo").wait_for(state="visible", timeout=30_000)
    komet.fill_order_dates(page, from_date, until_date)
    komet.search_orders(page)


def find_filtered_order(page, key: str) -> dict[str, str] | None:
    seen_pages: set[tuple[str, ...]] = set()
    for _page_number in range(1, 101):
        orders = komet.collect_order_rows(page)
        signature = komet.page_signature(orders)
        if not orders or signature in seen_pages:
            break
        seen_pages.add(signature)
        match = next((order for order in orders if komet.order_key(order) == key), None)
        if match:
            return match
        if not komet.click_next_page(page):
            break
    return None


def capture_download_error(page, index: int, order: dict[str, str], attempt: int) -> None:
    destination = ARTIFACTS_DIR / f"{index:03d}_{komet.safe_filename(order['order'])}_attempt{attempt}.png"
    try:
        page.screenshot(path=str(destination), full_page=True, timeout=10_000)
        print(f"Diagnóstico POSCO: captura {destination.name}; URL: {page.url}", flush=True)
    except Exception as error:
        print(f"Diagnóstico POSCO: no se pudo guardar captura ({error}); URL: {page.url}", flush=True)


def download_with_recovery(
    page, order: dict[str, str], index: int, from_date: date, until_date: date
) -> Path:
    key = komet.order_key(order)
    # The first successful mail flow performs a second search before downloads.
    # Resolve the row afresh because Komet can replace its grid after a search.
    current = find_filtered_order(page, key)
    if current is None:
        restore_filtered_orders(page, from_date, until_date)
        current = find_filtered_order(page, key)
    if current is None:
        raise RuntimeError(f"La orden {order['order']} ya no aparece en el rango filtrado.")
    try:
        return komet.download_export_details(page, current, index)
    except Exception as first_error:
        print(f"POSCO {order['order']}: primer intento de descarga falló: {first_error}", flush=True)
        capture_download_error(page, index, order, 1)
        restore_filtered_orders(page, from_date, until_date)
        current = find_filtered_order(page, key)
        if current is None:
            raise RuntimeError(f"No se pudo recuperar la orden {order['order']} tras el fallo de descarga.") from first_error
        try:
            return komet.download_export_details(page, current, index)
        except Exception as second_error:
            capture_download_error(page, index, order, 2)
            raise RuntimeError(
                f"La descarga de {order['order']} falló dos veces; URL actual: {page.url}. "
                f"Último error: {second_error}"
            ) from second_error


def transform_visible_orders(
    page, ledger: set[str], ignored: set[str], lookup: dict, template: bytes, report: dict,
    from_date: date, until_date: date, *, stage_mode: str = "off", stage_page=None,
    stage_state: dict[str, set[str]] | None = None,
) -> None:
    orders = collect_filtered_orders(page)
    for index, order in enumerate(orders, start=1):
        key = komet.order_key(order)
        if key in ignored:
            report["orders"].append({"order": order["order"], "key": key, "status": "manual"})
            print(f"POSCO {order['order']}: omitida por gestión manual.", flush=True)
            continue
        if stage_mode != "off":
            assert stage_state is not None
            stage_status = next((status for field, status in (
                ("legacy_orders", "legacy"), ("staged_orders", "already_staged"),
                ("needs_review", "needs_manual_review"),
            ) if key in stage_state[field]), None)
            if stage_status:
                report["orders"].append({"order": order["order"], "key": key, "status": stage_status})
                continue
        elif key in ledger:
            report["orders"].append({"order": order["order"], "key": key, "status": "already_transformed"})
            continue
        try:
            source = download_with_recovery(page, order, index, from_date, until_date)
            details, boxes = transform.parse_order_xls(source)
            verify_order_number(order["order"], details["order_number"])
            rows = transform.build_order_rows(details, boxes, lookup)
            destination = ARTIFACTS_DIR / f"{index:03d}_{komet.safe_filename(order['order'])}.xlsx"
            transform.write_order_workbook(template, rows, destination)
            if key not in ledger:
                transform.save_transformed_keys(ledger | {key}, LEDGER_PATH)
                ledger.add(key)
            if stage_mode != "off":
                stage_state["needs_review"].add(key)
                posco_import.save_stage_state(STAGE_LEDGER_PATH, stage_state)
                review_image = ARTIFACTS_DIR / f"{index:03d}_{komet.safe_filename(order['order'])}_posco_review.png"
                posco_import.stage_workbook(
                    stage_page, destination, order["order"], review_image,
                )
                stage_state["needs_review"].remove(key)
                stage_state["staged_orders"].add(key)
                posco_import.save_stage_state(STAGE_LEDGER_PATH, stage_state)
                report["orders"].append({
                    "order": order["order"], "key": key, "status": "staged_for_review",
                    "boxes": len(rows), "file": destination.name, "review_image": review_image.name,
                })
                print(f"POSCO {order['order']}: archivo cargado para revisión; Actualizar no se pulsó.", flush=True)
                continue
            report["orders"].append({
                "order": order["order"], "key": key, "status": "generated",
                "boxes": len(rows), "file": destination.name,
            })
            print(f"POSCO {order['order']}: {len(rows)} caja(s) en {destination.name}.", flush=True)
        except Exception as error:
            report["orders"].append({
                "order": order["order"], "key": key, "status": "error", "error": str(error),
            })
            print(f"ERROR POSCO {order['order']}: {error}", flush=True)
            if stage_mode == "all" and stage_state is not None and key in stage_state["needs_review"]:
                print("POSCO: carga sin confirmar; se detienen las siguientes para revisión manual.", flush=True)
                break


def run() -> dict:
    from_date, until_date = komet.komet_dates()
    stage_mode = os.environ.get("KOMET_POSCO_STAGE_MODE", "off").strip().lower()
    if stage_mode not in {"off", "all"}:
        raise ValueError("KOMET_POSCO_STAGE_MODE debe ser off o all.")
    report = {
        "from": from_date.isoformat(), "until": until_date.isoformat(),
        "orders": [], "global_error": None,
    }
    ARTIFACTS_DIR.mkdir(parents=True, exist_ok=True)
    try:
        ledger, ignored = transform.load_order_state(LEDGER_PATH)
        stage_state = posco_import.load_stage_state(STAGE_LEDGER_PATH) if stage_mode != "off" else None
        lookup = transform.load_homologation(transform.secret_workbook("KOMET_POSCO_MAPPING_B64"))
        template = transform.secret_workbook("KOMET_POSCO_TEMPLATE_B64")
        transform.validate_template(template)
        user = komet.required_secret("KOMET_USER")
        password = komet.required_secret("KOMET_PASSWORD")
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True)
            context = browser.new_context(viewport={"width": 1920, "height": 1080}, accept_downloads=True)
            page = context.new_page()
            try:
                komet.login_kometsales(page, user, password)
                komet.open_orders(page)
                komet.fill_order_dates(page, from_date, until_date)
                komet.search_orders(page)
                # Match the proven second filter/search pass in the mail flow.
                komet.fill_order_dates(page, from_date, until_date)
                komet.search_orders(page)
                stage_page = None
                if stage_mode != "off":
                    stage_page = context.new_page()
                    posco_import.open_import_screen(
                        stage_page, komet.required_secret("POSCO_USER"), komet.required_secret("POSCO_PASSWORD")
                    )
                transform_visible_orders(
                    page, ledger, ignored, lookup, template, report, from_date, until_date,
                    stage_mode=stage_mode, stage_page=stage_page, stage_state=stage_state,
                )
            finally:
                context.close()
                browser.close()
    except Exception as error:
        report["global_error"] = str(error)
        print(f"ERROR etapa POSCO: {error}", flush=True)
    finally:
        report["generated"] = sum(item["status"] == "generated" for item in report["orders"])
        report["errors"] = sum(item["status"] == "error" for item in report["orders"])
        report["manual"] = sum(item["status"] == "manual" for item in report["orders"])
        report["staged_for_review"] = sum(item["status"] == "staged_for_review" for item in report["orders"])
        REPORT_PATH.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(
            f"POSCO: {report['generated']} archivo(s), {report['errors']} error(es), "
            f"{sum(item['status'] == 'already_transformed' for item in report['orders'])} ya transformadas, "
            f"{report['manual']} de gestión manual.",
            flush=True,
        )
    return report


if __name__ == "__main__":
    result = run()
    if result["global_error"] or result["errors"]:
        raise SystemExit(1)

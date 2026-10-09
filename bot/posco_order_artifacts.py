"""Independent, review-only artifact pass after the Facturas Komet job."""

from __future__ import annotations

import json
import os
import re
from pathlib import Path

from playwright.sync_api import sync_playwright

if __package__:
    from . import facturas_komet as komet
    from . import posco_order_transform as transform
else:
    import facturas_komet as komet
    import posco_order_transform as transform


ARTIFACTS_DIR = Path("artifacts/facturas_komet/posco_orders")
REPORT_PATH = ARTIFACTS_DIR / "report.json"
LEDGER_PATH = Path(os.environ.get("KOMET_POSCO_TRANSFORMED_PATH", "bot/data/posco_transformados.json"))


def verify_order_number(selected: str, downloaded: str) -> None:
    selected_number = re.search(r"\d{6}", selected)
    downloaded_number = re.search(r"\d{6}", downloaded) or re.fullmatch(r"\d{1,6}", downloaded)
    if not selected_number or not downloaded_number or selected_number.group() != downloaded_number.group().zfill(6):
        raise transform.TransformationError(f"El XLS descargado es de la orden {downloaded}, no de {selected}.")


def transform_visible_orders(page, ledger: set[str], lookup: dict, template: bytes, report: dict) -> None:
    seen_pages: set[tuple[str, ...]] = set()
    index = 0
    for _page_number in range(1, 101):
        orders = komet.collect_order_rows(page)
        signature = komet.page_signature(orders)
        if not orders or signature in seen_pages:
            break
        seen_pages.add(signature)
        for order in orders:
            index += 1
            key = komet.order_key(order)
            if key in ledger:
                report["orders"].append({"order": order["order"], "key": key, "status": "already_transformed"})
                continue
            try:
                source = komet.download_export_details(page, order, index)
                details, boxes = transform.parse_order_xls(source)
                verify_order_number(order["order"], details["order_number"])
                rows = transform.build_order_rows(details, boxes, lookup)
                destination = ARTIFACTS_DIR / f"{index:03d}_{komet.safe_filename(order['order'])}.xlsx"
                transform.write_order_workbook(template, rows, destination)
                transform.save_transformed_keys(ledger | {key}, LEDGER_PATH)
                ledger.add(key)
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
        if not komet.click_next_page(page):
            break


def run() -> dict:
    from_date, until_date = komet.komet_dates()
    report = {
        "from": from_date.isoformat(), "until": until_date.isoformat(),
        "orders": [], "global_error": None,
    }
    ARTIFACTS_DIR.mkdir(parents=True, exist_ok=True)
    try:
        ledger = transform.load_transformed_keys(LEDGER_PATH)
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
                transform_visible_orders(page, ledger, lookup, template, report)
            finally:
                context.close()
                browser.close()
    except Exception as error:
        report["global_error"] = str(error)
        print(f"ERROR etapa POSCO: {error}", flush=True)
    finally:
        report["generated"] = sum(item["status"] == "generated" for item in report["orders"])
        report["errors"] = sum(item["status"] == "error" for item in report["orders"])
        REPORT_PATH.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(
            f"POSCO: {report['generated']} archivo(s), {report['errors']} error(es), "
            f"{sum(item['status'] == 'already_transformed' for item in report['orders'])} ya transformadas.",
            flush=True,
        )
    return report


if __name__ == "__main__":
    result = run()
    if result["global_error"]:
        raise SystemExit(1)

from __future__ import annotations

import json
import os
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from time import monotonic
from zoneinfo import ZoneInfo

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright


KOMET_LOGIN_URL = "https://app.kometsales.com/sign-in/login.do#st"
ARTIFACTS_DIR = Path("artifacts/facturas_komet")
CAPTURES_DIR = ARTIFACTS_DIR / "capturas"
SUMMARY_PATH = ARTIFACTS_DIR / "summary.json"
SENT_ORDERS_PATH = Path(
    os.environ.get("KOMET_SENT_ORDERS_PATH", "bot/data/facturas_enviadas.json")
)
DEFAULT_TIMEZONE = "America/Tijuana"
DEFAULT_TIMEOUT_MS = 60_000
ORDER_CODE_PATTERN = re.compile(r"\b(?:K2K\s*)?\d{6}\b|\b[A-Z]{1,3}\d{6}\b", re.I)
DATE_PATTERN = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")


def required_secret(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Falta configurar el secreto {name}.")
    return value


def visible_locator(locator):
    for index in range(locator.count()):
        candidate = locator.nth(index)
        try:
            if candidate.is_visible():
                return candidate
        except Exception:
            continue
    return None


def click_first_visible(page: Page, locators: list, description: str) -> None:
    for locator in locators:
        candidate = visible_locator(locator)
        if candidate is None:
            continue
        try:
            candidate.click(timeout=15_000)
            print(f"Clic: {description}", flush=True)
            return
        except PlaywrightTimeoutError:
            continue
    raise RuntimeError(f"No se encontró el control visible: {description}.")


def capture(page: Page, filename: str) -> None:
    CAPTURES_DIR.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(CAPTURES_DIR / filename), full_page=True)
    print(f"Captura guardada: {filename}", flush=True)


def wait_for_network(page: Page, timeout: int = DEFAULT_TIMEOUT_MS) -> None:
    try:
        page.wait_for_load_state("networkidle", timeout=timeout)
    except PlaywrightTimeoutError:
        print(f"Aviso: la página no llegó a networkidle. URL actual: {page.url}", flush=True)


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def safe_filename(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "-", value).strip("-.")
    return value[:80] or "orden"


def order_key(order: dict[str, str]) -> str:
    code = normalize_space(order.get("order", "")) or normalize_space(order.get("internal_id", ""))
    order_date = normalize_space(order.get("date", "")) or "sin-fecha"
    return f"{code}|{order_date}"


def load_sent_order_keys(path: Path = SENT_ORDERS_PATH) -> set[str]:
    if not path.exists():
        return set()
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise RuntimeError(f"No se pudo leer la bitácora de facturas enviadas: {path}.") from error
    keys = payload.get("sent_orders", [])
    if not isinstance(keys, list) or any(not isinstance(key, str) for key in keys):
        raise RuntimeError(f"La bitácora de facturas enviadas tiene un formato inválido: {path}.")
    return set(keys)


def save_sent_order_keys(keys: set[str], path: Path = SENT_ORDERS_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": 1,
        "sent_orders": sorted(keys),
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def komet_dates() -> tuple[date, date]:
    timezone_name = os.environ.get("KOMET_TIMEZONE", DEFAULT_TIMEZONE).strip() or DEFAULT_TIMEZONE
    today = datetime.now(ZoneInfo(timezone_name)).date()
    return today, today + timedelta(days=10)


def format_komet_date(value: date) -> str:
    # Komet muestra los filtros en formato MM/DD/YYYY.
    return value.strftime("%m/%d/%Y")


def login_kometsales(page: Page, user: str, password: str) -> None:
    print("Abriendo Kometsales...", flush=True)
    page.goto(KOMET_LOGIN_URL, wait_until="domcontentloaded", timeout=DEFAULT_TIMEOUT_MS)
    page.locator(
        'input[type="email"], input[placeholder*="usuario" i], input[name*="user" i], input[type="text"]'
    ).first.fill(user)
    page.locator(
        'input[type="password"], input[placeholder*="contraseña" i], input[placeholder*="password" i]'
    ).first.fill(password)
    click_first_visible(
        page,
        [
            page.get_by_role("button", name=re.compile(r"iniciar sesión|entrar|login", re.I)),
            page.locator('input[type="submit"]'),
        ],
        "Iniciar sesión",
    )
    wait_for_network(page)
    page.wait_for_timeout(2_000)

    account = page.get_by_text(re.compile(r"^\s*PACIFICA\s+FARMS\s*-\s*CAL\s*$", re.I))
    if account.count() > 0 and visible_locator(account) is not None:
        click_first_visible(page, [account], "PACIFICA FARMS - CAL")
        wait_for_network(page)
        page.wait_for_timeout(2_000)

    if "accounts.do" in page.url.lower():
        raise RuntimeError("No se pudo seleccionar PACIFICA FARMS - CAL en Kometsales.")
    print(f"Sesión de Kometsales iniciada. URL: {page.url}", flush=True)


def open_orders(page: Page) -> None:
    ventas_pattern = re.compile(r"^\s*ventas\s*$", re.I)
    click_first_visible(
        page,
        [
            page.get_by_role("button", name=ventas_pattern),
            page.get_by_role("link", name=ventas_pattern),
            page.get_by_text(ventas_pattern),
        ],
        "Ventas",
    )
    try:
        page.wait_for_url("**/orderSummary/list.do**", timeout=30_000)
    except PlaywrightTimeoutError:
        if "/orderSummary/list.do" not in page.url:
            raise RuntimeError(f"Se hizo clic en Ventas, pero Komet no abrió el resumen de órdenes. URL: {page.url}")
    wait_for_network(page)
    page.wait_for_timeout(1_500)
    if visible_locator(page.locator("#gridResults")) is None:
        raise RuntimeError("No se encontró la tabla de órdenes de Ventas en Kometsales.")
    capture(page, "01_ventas_seleccionada.png")


def fill_order_dates(page: Page, from_date: date, until_date: date) -> None:
    from_value = format_komet_date(from_date)
    until_value = format_komet_date(until_date)
    from_input = page.locator("#txtFromDateTo")
    until_input = page.locator("#txtDateTo")
    order_input = page.locator("#txtOrder")
    if from_input.count() != 1 or until_input.count() != 1:
        raise RuntimeError("No se encontraron exactamente los campos Orden desde y Orden hasta.")

    # No usar selectores genéricos como input[id*='to']: txtOrder también contiene "to".
    if order_input.count() == 1:
        order_input.fill("")
    from_input.fill(from_value)
    from_input.press("Tab")
    page.wait_for_timeout(250)
    until_input.fill(until_value)
    until_input.press("Tab")
    page.wait_for_timeout(250)

    actual_from = from_input.input_value()
    actual_until = until_input.input_value()
    actual_order = order_input.input_value() if order_input.count() == 1 else ""
    if actual_from != from_value or actual_until != until_value or actual_order:
        raise RuntimeError(
            "Komet no conservó el filtro solicitado: "
            f"Orden={actual_order!r}, desde={actual_from!r}, hasta={actual_until!r}."
        )
    print(f"Filtro validado: desde {actual_from} hasta {actual_until}.", flush=True)
    capture(page, "02_fechas_configuradas.png")


def search_orders(page: Page) -> None:
    pattern = re.compile(r"^\s*buscar\s*$", re.I)
    click_first_visible(
        page,
        [
            page.locator("#btnSearch"),
            page.get_by_role("button", name=pattern),
            page.get_by_role("link", name=pattern),
            page.get_by_text(pattern),
        ],
        "Buscar órdenes",
    )
    wait_for_network(page)
    page.wait_for_timeout(2_000)
    capture(page, "03_ordenes_filtradas.png")


def collect_order_rows(page: Page) -> list[dict[str, str]]:
    rows = []
    for row_index in range(page.locator("#gridResults tr").count()):
        row = page.locator("#gridResults tr").nth(row_index)
        checkbox = row.locator("input[id^='jqg_gridResults_']").first
        if checkbox.count() == 0:
            continue
        checkbox_id = checkbox.get_attribute("id") or ""
        match = re.search(r"jqg_gridResults_(.+)$", checkbox_id)
        if not match:
            continue
        internal_id = match.group(1)
        text = normalize_space(row.inner_text())
        order_match = ORDER_CODE_PATTERN.search(text)
        date_match = DATE_PATTERN.search(text)
        rows.append(
            {
                "internal_id": internal_id,
                "checkbox_id": checkbox_id,
                "order": normalize_space(order_match.group(0)) if order_match else internal_id,
                "date": date_match.group(0) if date_match else "",
            }
        )
    return rows


def page_signature(rows: list[dict[str, str]]) -> tuple[str, ...]:
    return tuple(row["internal_id"] for row in rows)


def click_next_page(page: Page) -> bool:
    candidates = [
        page.locator("#next_gridResults"),
        page.locator("#pager_gridResults #next_gridResults"),
        page.locator("a[title*='next' i], button[title*='next' i]"),
    ]
    for locator in candidates:
        candidate = visible_locator(locator)
        if candidate is None:
            continue
        classes = (candidate.get_attribute("class") or "").lower()
        aria_disabled = (candidate.get_attribute("aria-disabled") or "").lower()
        if "disabled" in classes or aria_disabled == "true":
            return False
        before = page_signature(collect_order_rows(page))
        candidate.click(timeout=15_000)
        deadline = monotonic() + 15
        while monotonic() < deadline:
            page.wait_for_timeout(500)
            after = page_signature(collect_order_rows(page))
            if after and after != before:
                wait_for_network(page, timeout=10_000)
                return True
        return True
    return False


def visible_dialog(page: Page):
    return visible_locator(page.locator(".ui-dialog:visible, [role='dialog']:visible"))


def check_document_boxes(page: Page, dialog) -> None:
    # En Komet las tres casillas aparecen en orden: Factura, Pick Ticket y Etiquetas.
    checkboxes = (dialog or page).locator('input[type="checkbox"]')
    if checkboxes.count() < 3:
        checkboxes = page.locator('input[type="checkbox"]:visible')
    if checkboxes.count() < 3:
        raise RuntimeError("El diálogo no mostró las casillas de Factura, Pick Ticket y Etiquetas.")
    for index in range(3):
        checkbox = checkboxes.nth(index)
        if not checkbox.is_checked():
            checkbox.check(force=True)


def fill_email_dialog(page: Page) -> object:
    dialog = visible_dialog(page)
    email = required_secret("KOMET_INVOICE_EMAIL")
    field_ids = [
        "#txtDialogOrderMailTo",
        "#txtDialogPickTicketMailTo",
        "#txtDialogLabelsMailTo",
    ]
    for selector in field_ids:
        field = page.locator(selector)
        field.wait_for(state="visible", timeout=20_000)
        field.fill(email)
    check_document_boxes(page, dialog)
    return dialog


def cancel_email_dialog(page: Page, dialog) -> None:
    pattern = re.compile(r"^\s*cancelar\s*$", re.I)
    roots = [dialog, page] if dialog else [page]
    for root in roots:
        if root is None:
            continue
        try:
            click_first_visible(
                page,
                [
                    root.get_by_role("button", name=pattern),
                    root.get_by_role("link", name=pattern),
                    root.get_by_text(pattern),
                ],
                "Cancelar envío de prueba",
            )
            page.wait_for_timeout(500)
            return
        except RuntimeError:
            continue
    raise RuntimeError("No se encontró el botón Cancelar del diálogo de correo.")


def send_email_dialog(page: Page, dialog) -> None:
    pattern = re.compile(r"^\s*enviar\s*$", re.I)
    roots = [dialog, page] if dialog else [page]
    clicked = False
    for root in roots:
        if root is None:
            continue
        try:
            click_first_visible(
                page,
                [
                    root.get_by_role("button", name=pattern),
                    root.get_by_role("link", name=pattern),
                    root.get_by_text(pattern),
                ],
                "Enviar documentos por correo",
            )
            clicked = True
            break
        except RuntimeError:
            continue
    if not clicked:
        raise RuntimeError("No se encontró el botón Enviar del diálogo de correo.")
    page.wait_for_timeout(1_500)
    if dialog is not None and visible_locator(dialog) is not None:
        raise RuntimeError("Komet mantuvo abierto el diálogo después de presionar Enviar.")


def process_order(page: Page, order: dict[str, str], index: int, mode: str) -> dict[str, str]:
    internal_id = order["internal_id"]
    order_label = safe_filename(f"{order['order']}-{order['date'] or internal_id}")
    row = page.locator(f"#gridResults tr:has(#{order['checkbox_id']})").first
    row.scroll_into_view_if_needed()
    row_checkbox = row.locator(f"#{order['checkbox_id']}")
    if not row_checkbox.is_checked():
        # Komet habilita las acciones de la orden después de marcar su cuadrito.
        row_checkbox.check(force=True)
    row.hover()

    action_span = page.locator(f".spnActions{internal_id}, #spnActions{internal_id}")
    more_actions = page.locator(f"#buttonContext{internal_id}")
    for target in [action_span, row.locator("td").last, row]:
        try:
            if target.count() > 0:
                target.hover()
                page.wait_for_timeout(250)
        except Exception:
            continue
        if visible_locator(more_actions) is not None:
            break
    more_actions.wait_for(state="visible", timeout=15_000)
    more_actions.click()
    capture(page, f"{index:03d}_{order_label}_01_acciones.png")

    email_menu = page.get_by_text(re.compile(r"^\s*Enviar Documentos? por Email\s*$", re.I))
    click_first_visible(page, [email_menu], "Enviar Documentos por Email")
    page.locator("#txtDialogOrderMailTo").wait_for(state="visible", timeout=20_000)
    dialog = fill_email_dialog(page)
    capture(page, f"{index:03d}_{order_label}_02_documentos_preparados.png")
    if mode == "cancel":
        cancel_email_dialog(page, dialog)
        capture(page, f"{index:03d}_{order_label}_03_cancelado.png")
        return {
            **order,
            "status": "cancelado",
            "mode": mode,
        }

    send_email_dialog(page, dialog)
    capture(page, f"{index:03d}_{order_label}_03_enviado.png")
    return {
        **order,
        "status": "enviado",
        "mode": mode,
    }


def run() -> None:
    mode = os.environ.get("KOMET_EMAIL_MODE", "cancel").strip().lower()
    if mode not in {"cancel", "send"}:
        raise RuntimeError("KOMET_EMAIL_MODE debe ser cancel o send.")
    if mode == "send" and os.environ.get("KOMET_ALLOW_SEND", "") != "YES":
        raise RuntimeError("El envío real está bloqueado. Define KOMET_ALLOW_SEND=YES sólo cuando sea autorizado.")

    user = required_secret("KOMET_USER")
    password = required_secret("KOMET_PASSWORD")
    from_date, until_date = komet_dates()
    sent_order_keys = load_sent_order_keys()
    summary = {
        "mode": mode,
        "timezone": os.environ.get("KOMET_TIMEZONE", DEFAULT_TIMEZONE),
        "orden_desde": format_komet_date(from_date),
        "orden_hasta": format_komet_date(until_date),
        "sent_orders_file": str(SENT_ORDERS_PATH),
        "orders": [],
    }
    ARTIFACTS_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1920, "height": 1080}, accept_downloads=True)
        page = context.new_page()
        try:
            login_kometsales(page, user, password)
            capture(page, "00_sesion_pacifica_farms.png")
            open_orders(page)
            fill_order_dates(page, from_date, until_date)
            search_orders(page)

            processed_pages: set[tuple[str, ...]] = set()
            next_index = 1
            for _page_number in range(1, 101):
                rows = collect_order_rows(page)
                if not rows or page_signature(rows) in processed_pages:
                    break
                processed_pages.add(page_signature(rows))
                print(f"Órdenes encontradas en la página: {len(rows)}", flush=True)
                for order in rows:
                    current_index = next_index
                    next_index += 1
                    key = order_key(order)
                    if key in sent_order_keys:
                        summary["orders"].append(
                            {
                                **order,
                                "order_key": key,
                                "status": "omitido_ya_enviado",
                                "mode": mode,
                            }
                        )
                        print(f"Orden {order['order']} omitida: ya fue enviada.", flush=True)
                        continue
                    try:
                        result = process_order(page, order, current_index, mode)
                        result["order_key"] = key
                        summary["orders"].append(result)
                        if result["status"] == "enviado":
                            sent_order_keys.add(key)
                            save_sent_order_keys(sent_order_keys)
                            print(f"Orden {order['order']} enviada y registrada.", flush=True)
                        else:
                            print(f"Orden {order['order']} preparada y cancelada.", flush=True)
                    except Exception as error:
                        summary["orders"].append(
                            {**order, "order_key": key, "status": "error", "error": str(error)}
                        )
                        capture(page, f"{current_index:03d}_{safe_filename(order['order'])}_99_error.png")
                        raise
                if not click_next_page(page):
                    break

            summary["total"] = len(summary["orders"])
            summary["cancelled"] = sum(item.get("status") == "cancelado" for item in summary["orders"])
            summary["sent"] = sum(item.get("status") == "enviado" for item in summary["orders"])
            summary["skipped"] = sum(item.get("status") == "omitido_ya_enviado" for item in summary["orders"])
            summary["errors"] = sum(item.get("status") == "error" for item in summary["orders"])
            print(
                f"Proceso terminado: {summary['sent']} enviadas, "
                f"{summary['cancelled']} canceladas y {summary['skipped']} omitidas.",
                flush=True,
            )
        except Exception:
            try:
                capture(page, "999_error_general.png")
            except Exception:
                pass
            raise
        finally:
            SUMMARY_PATH.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
            context.close()
            browser.close()


if __name__ == "__main__":
    run()

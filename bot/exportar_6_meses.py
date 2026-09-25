from __future__ import annotations

import os
import re
import calendar
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright


POSCO_URL = "http://3.132.9.174/Posco/"
ARTIFACTS_DIR = Path("artifacts/exportar_6_meses")
CAPTURES_DIR = ARTIFACTS_DIR / "capturas"
REPORTS_DIR = ARTIFACTS_DIR / "reportes"
BOT_TIMEZONE = ZoneInfo("America/Tijuana")


def required_secret(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Falta configurar el secreto {name}.")
    return value


def capture(page: Page, filename: str) -> None:
    CAPTURES_DIR.mkdir(parents=True, exist_ok=True)
    destination = CAPTURES_DIR / filename
    try:
        page.screenshot(path=str(destination), full_page=False, timeout=10_000)
    except PlaywrightTimeoutError:
        print(f"Aviso: no se pudo guardar la captura {filename}; el proceso continuará.", flush=True)
        return
    print(f"Captura guardada: {destination}", flush=True)


def click_visible_text(page: Page, text: str, description: str | None = None) -> None:
    pattern = re.compile(rf"^\s*{re.escape(text)}\s*$", re.IGNORECASE)
    candidates = [
        page.get_by_role("button", name=pattern),
        page.get_by_role("link", name=pattern),
        page.get_by_text(pattern),
    ]
    for candidate in candidates:
        for index in range(candidate.count()):
            element = candidate.nth(index)
            try:
                if element.is_visible():
                    element.click(timeout=10_000)
                    print(f"Clic: {description or text}", flush=True)
                    return
            except PlaywrightTimeoutError:
                continue
    raise RuntimeError(f"No se encontró el control visible: {description or text}.")


def open_orders_menu(page: Page) -> None:
    candidates = [
        page.get_by_role("link", name="Ordenes", exact=True),
        page.get_by_role("button", name="Ordenes", exact=True),
        page.locator('[ngbdropdowntoggle]:has-text("Ordenes")'),
        page.get_by_text("Ordenes", exact=True),
    ]
    for candidate in candidates:
        for index in range(candidate.count()):
            element = candidate.nth(index)
            try:
                if element.is_visible():
                    element.click(timeout=10_000)
                    return
            except PlaywrightTimeoutError:
                continue
    raise RuntimeError("No se encontró el menú superior Órdenes.")


def select_orders_option(page: Page) -> None:
    route = page.locator('a[href="#/list-orden-detalle"]')
    for index in range(route.count()):
        option = route.nth(index)
        if option.is_visible():
            option.click(timeout=10_000)
            return
    options = page.get_by_text("Ordenes", exact=True)
    for index in range(options.count()):
        option = options.nth(index)
        if option.is_visible() and option.get_attribute("href") == "#/list-orden-detalle":
            option.click(timeout=10_000)
            return
    raise RuntimeError("No se encontró la opción Órdenes dentro del menú.")


def select_active_status(page: Page) -> None:
    selects = page.locator("select")
    for index in range(selects.count()):
        select = selects.nth(index)
        labels = [text.strip() for text in select.locator("option").all_text_contents()]
        active_label = next((text for text in labels if text.casefold() == "activo"), None)
        has_all = any(text.casefold() in {"todos", "estatus todos", "all"} for text in labels)
        if active_label and has_all:
            select.select_option(label=active_label)
            print("Filtro de estado seleccionado: ACTIVO.", flush=True)
            return
    raise RuntimeError("No se encontró el filtro de estado con las opciones Todos y ACTIVO.")


def add_calendar_months(value: date, months: int) -> date:
    month_index = value.year * 12 + value.month - 1 + months
    year, zero_based_month = divmod(month_index, 12)
    month = zero_based_month + 1
    day = min(value.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def find_date_input(page: Page, label: str):
    group_input = page.locator(f'.input-group:has-text("{label}") input')
    if group_input.count() > 0:
        return group_input.first

    label_node = page.get_by_text(label, exact=True)
    if label_node.count() > 0:
        sibling_input = label_node.first.locator("xpath=..").locator("input")
        if sibling_input.count() > 0:
            return sibling_input.first

    raise RuntimeError(f"No se encontró el campo {label}.")


def set_load_date_mayor_six_months(page: Page) -> tuple[date, date]:
    """Keep Load Date Menor as POSCO set it; extend only Mayor to today + 6 months."""
    today = datetime.now(BOT_TIMEZONE).date()
    end_date = add_calendar_months(today, 6)
    mayor = find_date_input(page, "Load Date Mayor")
    input_type = (mayor.get_attribute("type") or "text").lower()
    formatted = end_date.isoformat() if input_type == "date" else end_date.strftime("%m/%d/%Y")
    mayor.fill(formatted)
    mayor.press("Tab")
    print(
        f"Load Date Menor se conserva; Load Date Mayor={end_date.isoformat()} "
        f"(hoy {today.isoformat()} + 6 meses calendario).",
        flush=True,
    )
    return today, end_date


def click_search(page: Page) -> None:
    candidates = [
        page.locator('button:has-text("Buscar"), a:has-text("Buscar")'),
        page.locator('[class*="btn"]:has-text("Buscar")'),
        page.get_by_text("Buscar", exact=True),
    ]
    for candidate in candidates:
        for index in range(candidate.count()):
            element = candidate.nth(index)
            if element.is_visible():
                element.click(timeout=10_000)
                print("Clic: Buscar", flush=True)
                return
    raise RuntimeError("No se encontró el botón Buscar en la página de Órdenes.")


def export_color_filter(page: Page) -> Path:
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    click_visible_text(page, "Exportar")
    page.wait_for_timeout(500)
    capture(page, "08_menu_exportar.png")

    with page.expect_download(timeout=60_000) as download_info:
        click_visible_text(page, "Exportar Color filtro")

    download = download_info.value
    filename = Path(download.suggested_filename).name or "exportar_color_filtro.xlsx"
    if Path(filename).suffix.lower() not in {".xlsx", ".xls", ".xlsm"}:
        filename = f"{Path(filename).stem or 'exportar_color_filtro'}.xlsx"
    destination = REPORTS_DIR / filename
    download.save_as(str(destination))
    print(f"Reporte descargado: {destination}", flush=True)
    return destination


def run() -> None:
    user = required_secret("POSCO_USER")
    password = required_secret("POSCO_PASSWORD")
    CAPTURES_DIR.mkdir(parents=True, exist_ok=True)
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1920, "height": 1080}, accept_downloads=True)
        page = context.new_page()
        try:
            print("Abriendo POSCO...", flush=True)
            page.goto(POSCO_URL, wait_until="domcontentloaded", timeout=60_000)
            capture(page, "01_login.png")
            page.locator('input[placeholder*="usuario@email.com" i], input[type="text"]').first.fill(user)
            page.locator('input[placeholder*="Password" i], input[type="password"]').first.fill(password)

            click_visible_text(page, "Iniciar Sesión")
            try:
                page.wait_for_load_state("networkidle", timeout=60_000)
            except PlaywrightTimeoutError:
                print("Aviso: POSCO continúa cargando después del inicio de sesión.", flush=True)
            page.wait_for_timeout(2_000)
            capture(page, "02_inicio_sesion.png")

            print("Abriendo Órdenes...", flush=True)
            open_orders_menu(page)
            page.wait_for_timeout(500)
            capture(page, "03_menu_ordenes.png")
            select_orders_option(page)
            try:
                page.wait_for_url("**/#/list-orden-detalle", timeout=30_000)
            except PlaywrightTimeoutError:
                if "list-orden-detalle" not in page.url:
                    raise RuntimeError(f"POSCO no abrió la lista de Órdenes. URL actual: {page.url}")
            page.wait_for_timeout(3_000)
            capture(page, "04_fechas_originales_posco.png")

            print("Conservando Load Date Menor y extendiendo Load Date Mayor a seis meses...", flush=True)
            today, end_date = set_load_date_mayor_six_months(page)
            capture(page, "05_rango_hasta_seis_meses.png")
            click_search(page)
            try:
                page.wait_for_load_state("networkidle", timeout=60_000)
            except PlaywrightTimeoutError:
                print("Aviso: POSCO sigue cargando resultados; se continuará con el filtro ACTIVO.", flush=True)
            page.wait_for_timeout(3_000)
            capture(page, "06_busqueda_rango_completada.png")

            print("Cambiando estado a ACTIVO después de buscar el rango...", flush=True)
            select_active_status(page)
            page.wait_for_timeout(6_000)
            capture(page, "07_ordenes_estado_activo.png")

            print("Exportando Exportar Color filtro...", flush=True)
            report = export_color_filter(page)
            capture(page, "09_exportacion_completada.png")
            print(
                f"EXPORTAR_6_MESES_OK inicio_posco_sin_cambios=true "
                f"hoy={today.isoformat()} load_date_mayor={end_date.isoformat()} reporte={report}",
                flush=True,
            )
        except Exception:
            capture(page, "99_error.png")
            raise
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    run()

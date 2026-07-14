import os
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


POSCO_URL = "http://3.132.9.174/Posco/"
CAPTURES_DIR = Path("artifacts/capturas")


def capture(page, name: str) -> None:
    """Guarda evidencia visual de cada paso para GitHub Actions."""
    CAPTURES_DIR.mkdir(parents=True, exist_ok=True)
    destination = CAPTURES_DIR / name
    page.screenshot(path=str(destination), full_page=True)
    print(f"Captura guardada: {destination}")


def required_secret(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Falta configurar el secreto {name}.")
    return value


def open_orders_menu(page) -> None:
    """Abre el menú superior Órdenes usando selectores tolerantes al HTML actual."""
    candidates = [
        page.get_by_role("link", name="Ordenes", exact=True),
        page.get_by_role("button", name="Ordenes", exact=True),
        page.locator('[ngbdropdowntoggle]:has-text("Ordenes")'),
        page.get_by_text("Ordenes", exact=True),
    ]

    for candidate in candidates:
        for index in range(candidate.count()):
            element = candidate.nth(index)
            if element.is_visible():
                element.click(timeout=10_000)
                return
    raise RuntimeError("No se encontró el menú superior Órdenes.")


def select_orders_option(page) -> None:
    """Selecciona la primera opción Órdenes dentro del menú desplegado."""
    exact_route = page.locator('a[href="#/list-orden-detalle"]')
    for index in range(exact_route.count()):
        option = exact_route.nth(index)
        if option.is_visible():
            option.click(timeout=10_000)
            return

    # Respaldo para implementaciones que no usan la clase Bootstrap dropdown-menu.
    options = page.get_by_text("Ordenes", exact=True)
    for index in range(options.count()):
        option = options.nth(index)
        if option.is_visible() and option.get_attribute("href") == "#/list-orden-detalle":
            option.click(timeout=10_000)
            return
    raise RuntimeError("No se encontró la opción Órdenes dentro del menú.")


def select_active_status(page) -> None:
    """Encuentra el filtro de estatus por sus opciones y selecciona ACTIVO."""
    selects = page.locator("select")
    for index in range(selects.count()):
        select = selects.nth(index)
        options = [text.strip() for text in select.locator("option").all_text_contents()]
        if "ACTIVO" in options:
            select.select_option(label="ACTIVO")
            return
    raise RuntimeError("No se encontró el filtro de estatus con la opción ACTIVO.")


def calculate_date_range(today: date | None = None) -> tuple[date, date]:
    """Devuelve el viernes anterior y el viernes tres semanas después."""
    current_date = today or datetime.now(ZoneInfo("America/Tijuana")).date()
    days_since_friday = (current_date.weekday() - 4) % 7
    if days_since_friday == 0:
        days_since_friday = 7
    previous_friday = current_date - timedelta(days=days_since_friday)
    return previous_friday, previous_friday + timedelta(weeks=3)


def find_date_input(page, label: str, fallback_index: int):
    group_input = page.locator(f'.input-group:has-text("{label}") input')
    if group_input.count() > 0:
        return group_input.first

    label_node = page.get_by_text(label, exact=True)
    if label_node.count() > 0:
        sibling_input = label_node.first.locator("xpath=..").locator("input")
        if sibling_input.count() > 0:
            return sibling_input.first

    date_inputs = page.locator('input[type="date"]')
    if date_inputs.count() > fallback_index:
        return date_inputs.nth(fallback_index)
    raise RuntimeError(f"No se encontró el campo {label}.")


def fill_date_input(locator, value: date) -> None:
    input_type = (locator.get_attribute("type") or "text").lower()
    formatted = value.isoformat() if input_type == "date" else value.strftime("%m/%d/%Y")
    locator.fill(formatted)
    locator.press("Tab")


def set_load_date_range(page) -> tuple[date, date]:
    start_date, end_date = calculate_date_range()
    menor = find_date_input(page, "Load Date Menor", 0)
    mayor = find_date_input(page, "Load Date Mayor", 1)
    fill_date_input(menor, start_date)
    fill_date_input(mayor, end_date)
    print(f"Rango configurado: {start_date.isoformat()} -> {end_date.isoformat()}")
    return start_date, end_date


def run() -> None:
    user = required_secret("POSCO_USER")
    password = required_secret("POSCO_PASSWORD")
    CAPTURES_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page = context.new_page()

        try:
            print("Abriendo Posco...")
            page.goto(POSCO_URL, wait_until="networkidle", timeout=60_000)
            capture(page, "01_login.png")

            print("Iniciando sesión...")
            page.locator('input[placeholder*="usuario@email.com" i], input[type="text"]').first.fill(user)
            page.locator('input[placeholder*="Password" i], input[type="password"]').first.fill(password)
            page.get_by_role("button", name="Iniciar Sesión").click(timeout=15_000)

            page.wait_for_load_state("networkidle", timeout=60_000)
            page.wait_for_timeout(2_000)
            capture(page, "02_dashboard.png")

            print("Abriendo menú Órdenes...")
            open_orders_menu(page)
            page.wait_for_timeout(800)
            capture(page, "03_menu_ordenes.png")

            print("Seleccionando Órdenes...")
            select_orders_option(page)
            try:
                page.wait_for_url("**/#/list-orden-detalle", timeout=30_000)
            except PlaywrightTimeoutError:
                print(f"La URL no cambió al patrón esperado. URL actual: {page.url}")
            page.wait_for_timeout(3_000)
            print("Configurando Load Date Menor y Load Date Mayor...")
            set_load_date_range(page)
            page.get_by_role("button", name="Buscar", exact=True).click(timeout=10_000)
            print("Esperando 30 segundos para que cargue el rango de fechas...")
            page.wait_for_timeout(30_000)
            capture(page, "04_rango_fechas.png")

            print("Cambiando el filtro de estatus a ACTIVO...")
            select_active_status(page)
            print("Esperando 15 segundos para que se aplique el filtro ACTIVO...")
            page.wait_for_timeout(15_000)
            capture(page, "05_status_activo.png")

            print(f"Paso exploratorio completado. URL final: {page.url}")
        except Exception:
            capture(page, "99_error.png")
            raise
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    run()

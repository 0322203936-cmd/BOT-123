import os
from pathlib import Path

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
        page.locator('a.dropdown-toggle:has-text("Ordenes")'),
        page.locator('a:has-text("Ordenes")'),
    ]

    for candidate in candidates:
        if candidate.count() > 0:
            candidate.first.click(timeout=10_000)
            return
    raise RuntimeError("No se encontró el menú superior Órdenes.")


def select_orders_option(page) -> None:
    """Selecciona la primera opción Órdenes dentro del menú desplegado."""
    visible_menu = page.locator(".dropdown-menu:visible")
    if visible_menu.count() > 0:
        option = visible_menu.get_by_text("Ordenes", exact=True)
        if option.count() > 0:
            option.first.click(timeout=10_000)
            return

    # Respaldo para implementaciones que no usan la clase Bootstrap dropdown-menu.
    options = page.get_by_text("Ordenes", exact=True)
    if options.count() >= 2:
        options.nth(1).click(timeout=10_000)
        return
    raise RuntimeError("No se encontró la opción Órdenes dentro del menú.")


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
                page.wait_for_url("**/#/revisar-ordenes", timeout=30_000)
            except PlaywrightTimeoutError:
                print(f"La URL no cambió al patrón esperado. URL actual: {page.url}")
            page.wait_for_timeout(3_000)
            capture(page, "04_ordenes_final.png")

            print(f"Paso exploratorio completado. URL final: {page.url}")
        except Exception:
            capture(page, "99_error.png")
            raise
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    run()

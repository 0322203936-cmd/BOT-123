"""Run the existing Excel Online "Ejecutar Generar Orden" button.

This is intentionally a separate browser bot. It does not edit the workbook,
does not press POSCO controls, and does not use Power Automate. Its only job is
to open the same SharePoint workbook the user opens, click the existing Office
Script button, and preserve the resulting PDF as a GitHub Actions artifact.
"""

from __future__ import annotations

import base64
import json
import os
import re
import time
from pathlib import Path

from playwright.sync_api import Browser, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright


DEFAULT_WORKBOOK_URL = (
    "https://pacificafarms.sharepoint.com/:x:/r/sites/requerimientovsproyeccion/"
    "_layouts/15/Doc.aspx?sourcedoc=%7B7FEB8241-AC2E-4964-B671-C4FAFF942F33%7D&"
    "file=PEDIDO%20Suzy%20-%20Mex%20y%20Can%20USA%2008-18-2026.xlsm&action=edit&mobileredirect=true"
)
BUTTON_NAME = "Ejecutar Generar Orden"
ARTIFACT_DIR = Path(os.environ.get("WHOLESALE_ORDER_ARTIFACT_DIR", "artifacts/wholesale_generar_orden"))


def required_secret(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(
            f"Falta configurar el secreto {name}. "
            "El bot necesita una cuenta de SharePoint para abrir Excel Online."
        )
    return value


def save_screenshot(page: Page, name: str) -> None:
    try:
        page.screenshot(path=str(ARTIFACT_DIR / name), full_page=True)
    except Exception as error:
        print(f"Aviso: no se pudo guardar {name}: {error}", flush=True)


def first_visible(page: Page, selectors: list[str], timeout: int = 5_000):
    for selector in selectors:
        locator = page.locator(selector).first
        try:
            locator.wait_for(state="visible", timeout=timeout)
            return locator
        except PlaywrightTimeoutError:
            continue
    return None


def load_storage_state(browser: Browser):
    encoded = os.environ.get("SHAREPOINT_STORAGE_STATE_B64", "").strip()
    if not encoded:
        return None
    try:
        return json.loads(base64.b64decode(encoded).decode("utf-8"))
    except (ValueError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RuntimeError("SHAREPOINT_STORAGE_STATE_B64 no contiene un estado válido de Playwright.") from error


def sign_in(page: Page) -> None:
    page.goto(
        os.environ.get("SHAREPOINT_WORKBOOK_URL", DEFAULT_WORKBOOK_URL),
        wait_until="domcontentloaded",
        timeout=120_000,
    )
    page.wait_for_timeout(4_000)

    if "login.microsoftonline.com" not in page.url.lower():
        return

    user = required_secret("SHAREPOINT_USER")
    password = required_secret("SHAREPOINT_PASSWORD")

    username = first_visible(
        page,
        ["input[name='loginfmt']", "input[type='email']", "input[name='loginfmt']"],
        timeout=15_000,
    )
    if username:
        username.fill(user)
        next_button = first_visible(page, ["input[type='submit']", "button:has-text('Siguiente')", "button:has-text('Next')"])
        if not next_button:
            raise RuntimeError("SharePoint pidió el usuario, pero no apareció el botón Siguiente.")
        next_button.click()
        page.wait_for_timeout(2_000)

    password_box = first_visible(
        page,
        ["input[name='passwd']", "input[type='password']"],
        timeout=15_000,
    )
    if not password_box:
        raise RuntimeError(
            "SharePoint no mostró el campo de contraseña. "
            "Si apareció MFA o una aprobación interactiva, este bot necesita una cuenta de automatización sin ese reto."
        )
    password_box.fill(password)
    submit = first_visible(
        page,
        ["input[type='submit']", "button:has-text('Iniciar sesión')", "button:has-text('Sign in')"],
        timeout=10_000,
    )
    if not submit:
        raise RuntimeError("SharePoint pidió la contraseña, pero no apareció el botón de inicio de sesión.")
    submit.click()
    page.wait_for_timeout(4_000)

    stay_signed_in = first_visible(
        page,
        ["input[value='No']", "button:has-text('No')", "button:has-text('No')"],
        timeout=5_000,
    )
    if stay_signed_in:
        stay_signed_in.click()
        page.wait_for_timeout(3_000)

    if "login.microsoftonline.com" in page.url.lower():
        raise RuntimeError(
            "SharePoint no terminó el inicio de sesión. Revisa si la cuenta requiere MFA, código o aprobación interactiva."
        )


def wait_for_workbook(page: Page) -> None:
    page.wait_for_timeout(10_000)
    if "login.microsoftonline.com" in page.url.lower():
        raise RuntimeError("La sesión de SharePoint no quedó autenticada.")
    try:
        page.get_by_text("PO Nuevo", exact=True).wait_for(timeout=30_000)
    except PlaywrightTimeoutError:
        # Excel Online can render the workbook on a canvas; the button is the
        # stronger readiness signal and is checked below.
        pass


def click_generate_order(page: Page) -> Path:
    button = None
    candidates = [
        page.get_by_role("button", name=BUTTON_NAME, exact=True),
        page.get_by_text(BUTTON_NAME, exact=True),
        page.locator(f"text={BUTTON_NAME}"),
    ]
    for candidate in candidates:
        try:
            candidate.first.wait_for(state="visible", timeout=20_000)
            button = candidate.first
            break
        except PlaywrightTimeoutError:
            continue
    if button is None:
        raise RuntimeError(f"No se encontró el botón {BUTTON_NAME!r} dentro de Excel Online.")

    save_screenshot(page, "01-antes-generar-orden.png")
    try:
        with page.expect_download(timeout=240_000) as download_info:
            button.click()
        download = download_info.value
        pdf_path = ARTIFACT_DIR / "generar-orden.pdf"
        download.save_as(str(pdf_path))
        save_screenshot(page, "02-despues-generar-orden.png")
        return pdf_path
    except PlaywrightTimeoutError as error:
        save_screenshot(page, "02-error-generar-orden.png")
        raise RuntimeError(
            "El botón fue localizado, pero Excel Online no entregó el PDF dentro de 4 minutos."
        ) from error


def main() -> None:
    ARTIFACT_DIR.mkdir(parents=True, exist_ok=True)
    started = time.time()
    summary: dict[str, object] = {
        "button": BUTTON_NAME,
        "workbook": os.environ.get("SHAREPOINT_WORKBOOK_URL", DEFAULT_WORKBOOK_URL),
        "status": "error",
    }

    try:
        with sync_playwright() as playwright:
            storage_state = load_storage_state(playwright.chromium)
            browser = playwright.chromium.launch(headless=True)
            context = browser.new_context(
                storage_state=storage_state,
                accept_downloads=True,
                locale="es-MX",
                timezone_id="America/Tijuana",
            )
            page = context.new_page()
            try:
                sign_in(page)
                wait_for_workbook(page)
                pdf_path = click_generate_order(page)
                summary.update(
                    {
                        "status": "completed",
                        "pdf": pdf_path.name,
                        "elapsed_seconds": round(time.time() - started, 1),
                        "url": page.url,
                    }
                )
                print(f"PDF generado: {pdf_path}", flush=True)
            finally:
                context.close()
                browser.close()
    except Exception as error:
        summary.update({"error": str(error), "elapsed_seconds": round(time.time() - started, 1)})
        raise
    finally:
        (ARTIFACT_DIR / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )


if __name__ == "__main__":
    main()

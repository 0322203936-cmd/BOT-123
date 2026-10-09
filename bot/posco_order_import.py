"""Stage a generated order workbook in POSCO's review screen (never click Actualizar)."""

import json
from pathlib import Path


POSCO_URL = "http://3.132.9.174/Posco/"
FORMAT_VALUE = "pacifica2"  # Pacifica - Andres Terriquez


def load_stage_state(path: Path) -> dict[str, set[str]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if payload.get("version") != 1:
        raise ValueError("Versión inválida de la bitácora de cargas POSCO.")
    fields = ("legacy_orders", "staged_orders", "needs_review")
    if any(not isinstance(payload.get(field), list) or any(not isinstance(key, str) for key in payload[field]) for field in fields):
        raise ValueError("Formato inválido de la bitácora de cargas POSCO.")
    state = {field: set(payload[field]) for field in fields}
    if any(state[left] & state[right] for left, right in (("legacy_orders", "staged_orders"), ("legacy_orders", "needs_review"), ("staged_orders", "needs_review"))):
        raise ValueError("Una orden aparece en más de un estado de cargas POSCO.")
    return state


def save_stage_state(path: Path, state: dict[str, set[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps({"version": 1, **{field: sorted(keys) for field, keys in state.items()}}, indent=2) + "\n",
        encoding="utf-8",
    )
    temporary.replace(path)


def open_import_screen(page, user: str, password: str) -> None:
    page.goto(POSCO_URL, wait_until="domcontentloaded", timeout=60_000)
    page.locator('input[placeholder*="usuario@email.com" i], input[type="text"]').first.fill(user)
    page.locator('input[placeholder*="Password" i], input[type="password"]').first.fill(password)
    page.get_by_role("button", name="Iniciar Sesión").click()
    page.locator("#DropdownOrdenes").wait_for(state="visible", timeout=60_000)
    page.locator("#DropdownOrdenes").click()
    page.get_by_role("link", name="Importar Ordenes", exact=True).click()
    page.get_by_role("button", name="Revisar Archivo").wait_for(state="visible", timeout=30_000)


def stage_workbook(page, workbook: Path, order_number: str, screenshot: Path) -> None:
    """Upload one workbook for review, without applying it to the order list."""
    if not workbook.is_file():
        raise FileNotFoundError(workbook)
    page.get_by_role("button", name="Revisar Archivo").click()
    page.locator('input[type="file"][name="file"]').set_input_files(str(workbook))
    page.locator('select[name="upload_mode"]').select_option(FORMAT_VALUE)
    try:
        page.get_by_role("button", name="Upload", exact=True).click()
        page.get_by_role("heading", name="Import excel").wait_for(state="hidden", timeout=600_000)
        page.get_by_text("No elements found", exact=True).wait_for(state="hidden", timeout=600_000)
    except Exception as error:
        raise RuntimeError(f"POSCO no confirmó filas para revisar de la orden {order_number}.") from error
    finally:
        screenshot.parent.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(screenshot), full_page=True)
    # Upload only stages a review. The final "Actualizar" action is deliberately absent.

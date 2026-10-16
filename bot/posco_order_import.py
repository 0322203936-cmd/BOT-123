"""Stage and, when explicitly enabled, apply one generated order in POSCO."""

import json
import re
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
    if not isinstance(payload.get("applied_orders", []), list) or any(not isinstance(key, str) for key in payload.get("applied_orders", [])):
        raise ValueError("Formato inválido de las órdenes confirmadas en POSCO.")
    state = {field: set(payload[field]) for field in fields}
    if "applied_orders" in payload:
        state["applied_orders"] = set(payload["applied_orders"])
    all_fields = tuple(state)
    if any(state[left] & state[right] for index, left in enumerate(all_fields) for right in all_fields[index + 1:]):
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


def stage_workbook(page, workbook: Path, order_number: str, screenshot: Path) -> str:
    """Upload one workbook without applying it; accept POSCO's 'Sin Cambios' result."""
    if not workbook.is_file():
        raise FileNotFoundError(workbook)
    if isinstance(page.url, str) and "#/revisar-ordenes" not in page.url:
        page.goto(POSCO_URL + "#/revisar-ordenes", wait_until="domcontentloaded", timeout=60_000)
        page.get_by_role("button", name="Revisar Archivo").wait_for(state="visible", timeout=30_000)
    page.get_by_role("button", name="Revisar Archivo").click()
    page.locator('input[type="file"][name="file"]').set_input_files(str(workbook))
    page.locator('select[name="upload_mode"]').select_option(FORMAT_VALUE)
    try:
        page.get_by_role("button", name="Upload", exact=True).click()
        page.get_by_role("heading", name="Import excel").wait_for(state="hidden", timeout=600_000)
        result = "no_changes" if page.get_by_text("Sin Cambios", exact=True).is_visible() else "uploaded"
    except Exception as error:
        raise RuntimeError(f"POSCO no terminó la carga de la orden {order_number}.") from error
    finally:
        screenshot.parent.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(screenshot), full_page=True)
    return result


def apply_staged_workbook(page, expected_identifier: str, expected_rows: int, screenshot: Path) -> None:
    """Apply only the freshly uploaded order; navigation is POSCO's success callback."""
    if expected_rows < 1:
        raise ValueError("No hay renglones para actualizar en POSCO.")
    matching_rows = page.get_by_role("cell", name=expected_identifier, exact=True).count()
    if matching_rows != expected_rows:
        raise RuntimeError(
            f"La revisión POSCO no coincide con {expected_identifier}: "
            f"se esperaban {expected_rows} renglones y aparecen {matching_rows}."
        )
    selection_text = page.get_by_text(re.compile(r"\d+\s+ordenes seleccionadas", re.I)).inner_text()
    selected_match = re.search(r"\d+", selection_text)
    if not selected_match or int(selected_match.group()) != expected_rows:
        raise RuntimeError(f"Selección POSCO inesperada para {expected_identifier}: {selection_text}.")
    update = page.get_by_role("button", name="Actualizar", exact=True)
    if not update.is_enabled():
        raise RuntimeError(f"POSCO no habilitó Actualizar para {expected_identifier}.")
    try:
        update.click()
        # POSCO navigates here only in the success callback of batchUpdateJSON,
        # after showing "Ordenes Actualizadas". A click alone is not confirmation.
        page.wait_for_url(re.compile(r"/Posco/#/list-orden-detalle(?:$|[?])"), timeout=120_000)
    finally:
        try:
            screenshot.parent.mkdir(parents=True, exist_ok=True)
            page.screenshot(path=str(screenshot), full_page=True, timeout=15_000)
        except Exception as error:
            print(f"Aviso: no se pudo guardar evidencia final POSCO ({error}).", flush=True)

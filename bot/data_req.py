import os
import re
import math
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo
import pandas as pd

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from sharepoint_sync import graph_token, GRAPH_URL, resolve_sharepoint_item_by_url, graph_headers
from excel_range_sync import graph_request
from data_proy_refresh import rebuild_append1_output, refresh_weeks_pivot

POSCO_URL = "http://3.132.9.174/Posco/"
CAPTURES_DIR = Path("artifacts/capturas")
REPORTS_DIR = Path("artifacts/reportes")

REQ_PROY_URL = "https://pacificafarms.sharepoint.com/:x:/r/sites/requerimientovsproyeccion/_layouts/15/Doc.aspx?sourcedoc=%7BB6111299-1373-4717-A2B7-3D507AD77A8A%7D&file=Requerimiento%20vs%20proyeccion.xlsm&action=default&mobileredirect=true"
DATA_REQ_CHUNK_SIZE = 1000
EXCEL_DATE_EPOCH = datetime(1899, 12, 30)


def parse_excel_datetime(value):
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, str):
        candidate = value.strip()
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}(?:[T ][0-9:.+-]+)?", candidate):
            return None
        try:
            parsed = datetime.fromisoformat(candidate)
            return parsed.replace(tzinfo=None)
        except ValueError:
            return None
    return None


def excel_values_equal(expected, actual) -> bool:
    expected_blank = expected is None or expected == ""
    actual_blank = actual is None or actual == ""
    if expected_blank or actual_blank:
        return expected_blank and actual_blank

    expected_date = parse_excel_datetime(expected)
    if expected_date is not None:
        actual_date = parse_excel_datetime(actual)
        if actual_date is None and isinstance(actual, (int, float)):
            actual_date = EXCEL_DATE_EPOCH + timedelta(days=float(actual))
        if actual_date is not None:
            return expected_date == actual_date

    if not isinstance(expected, bool) and not isinstance(actual, bool):
        try:
            exp_num = float(expected)
            act_num = float(actual)
            return math.isclose(exp_num, act_num, rel_tol=1e-9, abs_tol=1e-9)
        except (ValueError, TypeError):
            pass

    return expected == actual


def excel_rows_equal(expected, actual) -> bool:
    width = max(len(expected), len(actual))
    expected_values = list(expected) + [""] * (width - len(expected))
    actual_values = list(actual) + [""] * (width - len(actual))
    return all(
        excel_values_equal(expected_value, actual_value)
        for expected_value, actual_value in zip(expected_values, actual_values)
    )


class ResilientWorkbookSession:
    def __init__(self, workbook_url, headers, request_func=graph_request):
        self.workbook_url = workbook_url
        self.base_headers = headers
        self.request_func = request_func
        self.session_headers = None

    def open(self):
        session = self.request_func(
            "POST",
            f"{self.workbook_url}/createSession",
            self.base_headers,
            json={"persistChanges": True},
            timeout=60,
        ).json()
        session_id = session.get("id")
        if not session_id:
            raise RuntimeError("Microsoft Graph no devolvio un identificador de sesion.")
        self.session_headers = {
            **self.base_headers,
            "workbook-session-id": session_id,
        }
        return self.session_headers

    @staticmethod
    def _session_is_invalid(error: Exception) -> bool:
        message = str(error).lower()
        return (
            "target session is invalid" in message
            or "invalidsession" in message
            or "invalid session" in message
        )

    def close(self):
        if not self.session_headers:
            return
        try:
            self.request_func(
                "POST",
                f"{self.workbook_url}/closeSession",
                self.session_headers,
                timeout=30,
            )
        except Exception:
            pass
        finally:
            self.session_headers = None

    def renew(self):
        self.close()
        print(
            "La sesion de Excel fue invalidada; abriendo una nueva "
            "y repitiendo el bloque pendiente...",
            flush=True,
        )
        self.open()

    def request(self, method, url, **kwargs):
        if not self.session_headers:
            self.open()
        try:
            return self.request_func(
                method,
                url,
                self.session_headers,
                **kwargs,
            )
        except RuntimeError as error:
            if not self._session_is_invalid(error):
                raise
            self.renew()
            return self.request_func(
                method,
                url,
                self.session_headers,
                **kwargs,
            )


def session_graph_request(session):
    def request(method, url, _headers, **kwargs):
        return session.request(method, url, **kwargs)

    return request


def refresh_data_req_outputs(
    workbook_url,
    headers,
    session_factory=None,
    sleep_func=time.sleep,
):
    factory = session_factory or ResilientWorkbookSession

    print("Abriendo una sesion nueva para actualizar Append1...")
    append_session = factory(workbook_url, headers)
    append_session.open()
    try:
        append_rows = rebuild_append1_output(
            session_graph_request(append_session),
            workbook_url,
            append_session.session_headers,
        )
        print(f"Append1 actualizada con exito: {append_rows} filas.")
    finally:
        append_session.close()

    print("Abriendo una sesion nueva para actualizar Weeks x FechaProduccion...")
    pivot_session = factory(workbook_url, headers)
    pivot_session.open()
    try:
        refresh_weeks_pivot(
            session_graph_request(pivot_session),
            workbook_url,
            pivot_session.session_headers,
        )
        print("Esperando que Excel Online guarde la tabla dinamica...")
        sleep_func(5)
        print(
            "Salidas de Data Req actualizadas con exito: "
            f"Append1={append_rows} filas; "
            "Weeks x FechaProduccion=actualizada."
        )
    finally:
        pivot_session.close()

    return append_rows


def env_flag(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "si", "sí"}

def capture(page, name: str) -> None:
    CAPTURES_DIR.mkdir(parents=True, exist_ok=True)
    destination = CAPTURES_DIR / name
    try:
        page.screenshot(path=str(destination), full_page=False)
    except Exception as e:
        print(f"Error al tomar captura {name}: {e}")
    print(f"Captura intentada: {destination}")

def required_secret(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Falta configurar el secreto {name}.")
    return value

def open_orders_menu(page) -> None:
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
    exact_route = page.locator('a[href="#/list-orden-detalle"]')
    for index in range(exact_route.count()):
        option = exact_route.nth(index)
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

def select_active_status(page) -> None:
    selects = page.locator("select")
    for index in range(selects.count()):
        select = selects.nth(index)
        options = [text.strip() for text in select.locator("option").all_text_contents()]
        if "ACTIVO" in options:
            select.select_option(label="ACTIVO")
            return
    raise RuntimeError("No se encontró el filtro de estatus con la opción ACTIVO.")

def calculate_date_range(today: date | None = None) -> tuple[date, date]:
    """Devuelve el viernes anterior y el viernes 12 semanas después para DataReq."""
    current_date = today or datetime.now(ZoneInfo("America/Tijuana")).date()
    days_since_friday = (current_date.weekday() - 4) % 7
    if days_since_friday == 0:
        days_since_friday = 7
    previous_friday = current_date - timedelta(days=days_since_friday)
    return previous_friday, previous_friday + timedelta(weeks=13)

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

def click_search(page) -> None:
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
                return
    raise RuntimeError("No se encontró el control Buscar.")

def click_visible_text(page, text: str) -> None:
    candidates = page.get_by_text(text, exact=True)
    for index in range(candidates.count()):
        element = candidates.nth(index)
        if element.is_visible():
            element.click(timeout=30_000)
            return
    raise RuntimeError(f"No se encontró el control visible {text}.")

def export_color_filter(page) -> Path:
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    click_visible_text(page, "Exportar")
    page.wait_for_timeout(500)
    capture(page, "06_menu_exportar.png")

    with page.expect_download(timeout=120_000) as download_info:
        click_visible_text(page, "Exportar Color filtro")

    download = download_info.value
    filename = Path(download.suggested_filename).name or "exportar_color_filtro.xlsx"
    destination = REPORTS_DIR / filename
    download.save_as(str(destination))
    print(f"Reporte descargado: {destination}")
    return destination

def clean_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, float) and not math.isfinite(value):
        return ""
    if isinstance(value, datetime) or isinstance(value, date):
        return value.isoformat()
    return value

def process_and_upload(report_path: Path):
    if os.environ.get("SHAREPOINT_UPLOAD", "true").lower() not in {"1", "true", "yes", "si", "sí"}:
        print("Modo prueba: SHAREPOINT_UPLOAD esta apagado.")
        return

    print("Procesando Excel en crudo...")
    # Leer sin encabezados para conservar todo tal cual, la fila 0 serán los títulos de Posco
    # Pero el usuario dijo "sin tocar titulos lo que pegara va ser lo del reporte de posco"
    # Wait, usually when we say "sin tocar titulos", it means don't touch the destination headers.
    # The source data: we should omit the headers of the Posco report, so we only paste DATA.
    # So we read normally (header=0) which uses row 0 as column names.
    df = pd.read_excel(report_path)
    
    # We need to take columns A to R (indices 0 to 17), but SKIP N (index 13).
    selected_cols = [i for i in range(18) if i != 13]
    
    # Extraer data (lista de listas), descartando los títulos del DataFrame
    data_subset = df.iloc[:, selected_cols]
    
    # Convertir a array limpio
    new_data = []
    for _, row in data_subset.iterrows():
        new_data.append([clean_value(x) for x in row.values])
        
    print(f"Se extraerán {len(new_data)} filas (cada una de {len(new_data[0]) if new_data else 0} columnas).")

    print("Obteniendo token de Microsoft Graph...")
    token = graph_token()
    item_req = resolve_sharepoint_item_by_url(token, REQ_PROY_URL)
    drive_id = item_req["parentReference"]["driveId"]
    workbook_url = f"{GRAPH_URL}/drives/{drive_id}/items/{item_req['id']}/workbook"
    headers = {**graph_headers(token), "Content-Type": "application/json"}
    
    print("Abriendo sesion persistente en DataReq...")
    session = ResilientWorkbookSession(workbook_url, headers)
    session.open()
    
    try:
        print("Buscando límite de datos previos en DataReq...")
        used = session.request("GET", f"{workbook_url}/worksheets/DataReq/usedRange", timeout=120).json()
        address = used.get("address", "")
        match = re.search(r'!([A-Za-z]+)\d+:([A-Za-z]+)(\d+)', address)
        end_row = int(match.group(3)) if match else 1000
        
        if end_row >= 2:
            print(f"Limpiando datos existentes B2:R{end_row} y Z2:Z{end_row}...")
            # Microsoft Graph clear endpoint
            session.request("POST", f"{workbook_url}/worksheets/DataReq/range(address='B2:R{end_row}')/clear",
                            json={"applyTo": "Contents"}, timeout=120)
            session.request("POST", f"{workbook_url}/worksheets/DataReq/range(address='Z2:Z{end_row}')/clear",
                            json={"applyTo": "Contents"}, timeout=120)
        
        if len(new_data) > 0:
            z_data = []
            for row in new_data:
                # El indice 13 corresponde a la columna O en nuestra proyección B..R
                z_data.append([row[13]])
                row[13] = ""
            
            for i in range(0, len(new_data), DATA_REQ_CHUNK_SIZE):
                chunk = new_data[i : i + DATA_REQ_CHUNK_SIZE]
                chunk_z = z_data[i : i + DATA_REQ_CHUNK_SIZE]
                chunk_start = 2 + i
                chunk_end = chunk_start + len(chunk) - 1
                
                chunk_addr = f"B{chunk_start}:R{chunk_end}"
                print(f"Pegando chunk {chunk_addr}...")
                session.request("PATCH", f"{workbook_url}/worksheets/DataReq/range(address='{chunk_addr}')",
                                json={"values": chunk}, timeout=120)
                              
                chunk_addr_z = f"Z{chunk_start}:Z{chunk_end}"
                print(f"Pegando chunk {chunk_addr_z}...")
                session.request("PATCH", f"{workbook_url}/worksheets/DataReq/range(address='{chunk_addr_z}')",
                                json={"values": chunk_z}, timeout=120)
                time.sleep(1)

            time.sleep(3)
            last_row = len(new_data) + 1
            last_main = session.request(
                "GET",
                f"{workbook_url}/worksheets/DataReq/range(address='B{last_row}:R{last_row}')",
                timeout=120,
            ).json().get("values", [[]])[0]
            last_z = session.request(
                "GET",
                f"{workbook_url}/worksheets/DataReq/range(address='Z{last_row}:Z{last_row}')",
                timeout=120,
            ).json().get("values", [[]])[0]
            if not excel_rows_equal(new_data[-1], last_main):
                print(f"ERROR DE VERIFICACION B:R.\nEsperado (Posco): {new_data[-1]}\nObtenido (SharePoint): {last_main}", flush=True)
                raise RuntimeError("La verificacion final de DataReq fallo en las columnas B:R.")
            if not excel_rows_equal(z_data[-1], last_z):
                print(f"ERROR DE VERIFICACION Z.\nEsperado (Posco): {z_data[-1]}\nObtenido (SharePoint): {last_z}", flush=True)
                raise RuntimeError("La verificacion final de DataReq fallo en la columna Z.")

            print("Datos copiados exitosamente a DataReq.")
    finally:
        session.close()

    refresh_data_req_outputs(workbook_url, headers)

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
                pass
            page.wait_for_timeout(3_000)
            print("Configurando Load Date Menor y Load Date Mayor (13 semanas)...")
            set_load_date_range(page)
            page.wait_for_timeout(1_000)
            click_search(page)
            print("Esperando 4 minutos para que cargue el rango de fechas...")
            page.wait_for_timeout(240_000)
            capture(page, "04_rango_fechas.png")

            print("Cambiando el filtro de estatus a ACTIVO...")
            select_active_status(page)
            print("Esperando 6 segundos para que se aplique el filtro ACTIVO...")
            page.wait_for_timeout(6_000)
            capture(page, "05_status_activo.png")

            print("Abriendo Exportar y descargando Exportar Color filtro...")
            downloaded_report = export_color_filter(page)
            
            # Procesar el reporte y pegarlo en DataReq
            process_and_upload(downloaded_report)
            
            page.wait_for_timeout(5_000)
            capture(page, "07_exportacion_completada.png")

        except Exception:
            capture(page, "99_error.png")
            raise
        finally:
            context.close()
            browser.close()

if __name__ == "__main__":
    run()

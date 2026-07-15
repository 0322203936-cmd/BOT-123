import base64
import hashlib
import os
import time
from copy import copy
from pathlib import Path
from zipfile import ZipFile

import requests
from openpyxl import load_workbook


GRAPH_URL = "https://graph.microsoft.com/v1.0"
SHAREPOINT_FILE_URL = (
    "https://pacificafarms.sharepoint.com/:x:/r/sites/"
    "requerimientovsproyeccion/_layouts/15/Doc.aspx?"
    "sourcedoc=%7BF5574BDC-EC82-44BF-8D2E-B42CEB29D586%7D&"
    "file=Reunion%201-2-3%20Test.xlsm&action=default&mobileredirect=true"
)
SHAREPOINT_DIR = Path("artifacts/sharepoint")


def required_secret(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Falta configurar el secreto {name}.")
    return value


def graph_token() -> str:
    tenant_id = required_secret("SHAREPOINT_TENANT_ID")
    response = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={
            "client_id": required_secret("SHAREPOINT_CLIENT_ID"),
            "client_secret": required_secret("SHAREPOINT_CLIENT_SECRET"),
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()["access_token"]


def graph_headers(token: str) -> dict[str, str]:
    return {"Authorization": f"Bearer {token}"}


def share_id(url: str) -> str:
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii")
    return f"u!{encoded.rstrip('=')}"


def resolve_sharepoint_item(token: str) -> dict:
    response = requests.get(
        f"{GRAPH_URL}/shares/{share_id(SHAREPOINT_FILE_URL)}/driveItem",
        headers=graph_headers(token),
        timeout=30,
    )
    response.raise_for_status()
    item = response.json()
    if item.get("name", "").lower() != "reunion 1-2-3 test.xlsm":
        raise RuntimeError(f"SharePoint devolvió un archivo inesperado: {item.get('name')}.")
    return item


def download_sharepoint_workbook(token: str, item: dict) -> Path:
    drive_id = item["parentReference"]["driveId"]
    item_id = item["id"]
    response = requests.get(
        f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/content",
        headers=graph_headers(token),
        timeout=120,
    )
    response.raise_for_status()
    SHAREPOINT_DIR.mkdir(parents=True, exist_ok=True)
    destination = SHAREPOINT_DIR / "Reunion 1-2-3 Test_original.xlsm"
    destination.write_bytes(response.content)
    print(f"Libro de SharePoint descargado: {destination}")
    return destination


def vba_hash(path: Path) -> str:
    with ZipFile(path) as archive:
        try:
            content = archive.read("xl/vbaProject.bin")
        except KeyError as exc:
            raise RuntimeError("El libro XLSM no contiene xl/vbaProject.bin.") from exc
    return hashlib.sha256(content).hexdigest()


def snapshot_columns_from_s(worksheet, last_column: int) -> list[list[tuple]]:
    return [
        [
            (
                worksheet.cell(row=row, column=column).value,
                worksheet.cell(row=row, column=column).data_type,
                worksheet.cell(row=row, column=column).number_format,
                copy(worksheet.cell(row=row, column=column)._style),
                worksheet.cell(row=row, column=column).hyperlink.target
                if worksheet.cell(row=row, column=column).hyperlink
                else None,
                worksheet.cell(row=row, column=column).comment.text
                if worksheet.cell(row=row, column=column).comment
                else None,
            )
            for column in range(19, last_column + 1)
        ]
        for row in range(1, worksheet.max_row + 1)
    ]


def clear_cell(cell) -> None:
    cell.value = None
    cell._style = None
    cell._hyperlink = None
    cell.comment = None


def copy_cell(source, destination) -> None:
    destination.value = source.value
    destination._style = copy(source._style)
    destination._hyperlink = copy(source.hyperlink)
    destination.comment = copy(source.comment)


def replace_pegar_data(report_path: Path, workbook_path: Path) -> Path:
    source_book = load_workbook(report_path, data_only=False)
    source = source_book.active
    if source.max_column != 18:
        raise RuntimeError(f"El reporte formateado debe tener 18 columnas; tiene {source.max_column}.")

    original_vba = vba_hash(workbook_path)
    target_book = load_workbook(workbook_path, keep_vba=True, keep_links=True)
    if "PegarData" not in target_book.sheetnames:
        raise RuntimeError("No existe la hoja PegarData en el libro de SharePoint.")
    target = target_book["PegarData"]

    preserved_rows = target.max_row
    preserved_last_column = target.max_column
    preserved_from_s = snapshot_columns_from_s(target, preserved_last_column)

    for row in range(1, target.max_row + 1):
        for column in range(1, 19):
            clear_cell(target.cell(row=row, column=column))

    for row in range(1, source.max_row + 1):
        for column in range(1, 19):
            copy_cell(
                source.cell(row=row, column=column),
                target.cell(row=row, column=column),
            )

    for column in range(1, 19):
        letter = target.cell(row=1, column=column).column_letter
        target.column_dimensions[letter].width = source.column_dimensions[letter].width
    for row in range(1, source.max_row + 1):
        target.row_dimensions[row].height = source.row_dimensions[row].height

    target.auto_filter.ref = f"A1:R{source.max_row}"
    destination = SHAREPOINT_DIR / "Reunion 1-2-3 Test_actualizado.xlsm"
    target_book.save(destination)
    target_book.close()

    if vba_hash(destination) != original_vba:
        raise RuntimeError("La validación detectó un cambio en el proyecto VBA.")

    verification_book = load_workbook(destination, keep_vba=True, keep_links=True, data_only=False)
    verification = verification_book["PegarData"]
    verification_from_s = snapshot_columns_from_s(verification, preserved_last_column)
    if verification_from_s[:preserved_rows] != preserved_from_s:
        raise RuntimeError("La validación detectó cambios desde la columna S.")
    for row in range(1, source.max_row + 1):
        for column in range(1, 19):
            if verification.cell(row=row, column=column).value != source.cell(
                row=row, column=column
            ).value:
                raise RuntimeError(f"El valor pegado no coincide en fila {row}, columna {column}.")
    verification_book.close()
    source_book.close()
    print(
        f"PegarData validada: A1:R{source.max_row} reemplazado; "
        "columnas desde S y VBA conservados."
    )
    return destination


def upload_sharepoint_workbook(token: str, item: dict, workbook_path: Path) -> None:
    drive_id = item["parentReference"]["driveId"]
    item_id = item["id"]
    content = workbook_path.read_bytes()
    upload_url = f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/content"
    headers = {
        **graph_headers(token),
        "Content-Type": "application/octet-stream",
        "If-Match": item["eTag"],
    }

    for attempt in range(1, 7):
        response = requests.put(upload_url, headers=headers, data=content, timeout=180)
        if response.status_code == 423 and attempt < 6:
            print(
                f"El libro está bloqueado en SharePoint. "
                f"Reintento {attempt}/5 en 20 segundos..."
            )
            time.sleep(20)
            continue
        if response.status_code == 412:
            raise RuntimeError(
                "El libro cambió en SharePoint durante el proceso; "
                "se canceló la subida para no sobrescribir cambios recientes."
            )
        response.raise_for_status()
        break

    verification_path = SHAREPOINT_DIR / "Reunion 1-2-3 Test_verificacion.xlsm"
    last_error = None
    for attempt in range(1, 6):
        verification = requests.get(
            upload_url,
            headers=graph_headers(token),
            timeout=180,
        )
        verification.raise_for_status()
        verification_path.write_bytes(verification.content)
        try:
            validate_uploaded_workbook(workbook_path, verification_path)
            print("Libro actualizado y verificado correctamente en SharePoint.")
            return
        except RuntimeError as exc:
            last_error = exc
            if attempt < 5:
                print(f"SharePoint aún procesa el archivo. Verificación {attempt}/5...")
                time.sleep(5)
    raise RuntimeError("No se pudo validar el archivo guardado en SharePoint.") from last_error


def validate_uploaded_workbook(expected_path: Path, remote_path: Path) -> None:
    if vba_hash(remote_path) != vba_hash(expected_path):
        raise RuntimeError("El VBA remoto no coincide con el archivo validado.")

    expected_book = load_workbook(expected_path, keep_vba=True, keep_links=True, data_only=False)
    remote_book = load_workbook(remote_path, keep_vba=True, keep_links=True, data_only=False)
    try:
        if remote_book.sheetnames != expected_book.sheetnames:
            raise RuntimeError("Las hojas del libro remoto no coinciden.")

        expected = expected_book["PegarData"]
        remote = remote_book["PegarData"]
        if (remote.max_row, remote.max_column) != (expected.max_row, expected.max_column):
            raise RuntimeError("Las dimensiones de PegarData no coinciden.")

        for row in range(1, expected.max_row + 1):
            for column in range(1, expected.max_column + 1):
                expected_cell = expected.cell(row=row, column=column)
                remote_cell = remote.cell(row=row, column=column)
                if (
                    remote_cell.value != expected_cell.value
                    or remote_cell.data_type != expected_cell.data_type
                    or remote_cell.number_format != expected_cell.number_format
                    or remote_cell._style != expected_cell._style
                ):
                    raise RuntimeError(
                        f"PegarData remota no coincide en fila {row}, columna {column}."
                    )
    finally:
        expected_book.close()
        remote_book.close()


def sync_report_to_sharepoint(report_path: Path, upload: bool) -> Path:
    token = graph_token()
    item = resolve_sharepoint_item(token)
    original = download_sharepoint_workbook(token, item)
    updated = replace_pegar_data(report_path, original)
    if upload:
        upload_sharepoint_workbook(token, item, updated)
    else:
        print("Modo de prueba: el libro fue validado, pero no se subió a SharePoint.")
    return updated

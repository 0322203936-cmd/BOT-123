"""Generate the Wholesale order PDF through Microsoft Graph.

The former implementation opened Excel Online and clicked the Office Script
button.  Graph cannot invoke an Office Script directly, so this bot reproduces
the script's observable behavior safely:

* download the source workbook through Graph;
* create a short-lived workbook copy with only ``PO Nuevo`` visible;
* apply the same print area, orientation, and letter paper settings;
* ask Graph to convert that temporary workbook to PDF; and
* delete the temporary SharePoint item in a ``finally`` block.

The permanent XLSM is never overwritten and no interactive SharePoint login is
needed.  The existing workbook, macros, and other bots are left untouched.
"""

from __future__ import annotations

import json
import os
import re
import tempfile
import time
from pathlib import Path
from urllib.parse import quote
from uuid import uuid4
from zipfile import ZipFile

import requests

from sharepoint_sync import (
    GRAPH_URL,
    graph_headers,
    graph_token,
    resolve_sharepoint_item_by_url,
)


DEFAULT_WORKBOOK_URL = (
    "https://pacificafarms.sharepoint.com/:x:/r/sites/requerimientovsproyeccion/"
    "_layouts/15/Doc.aspx?sourcedoc=%7B7FEB8241-AC2E-4964-B671-C4FAFF942F33%7D&"
    "file=PEDIDO%20Suzy%20-%20Mex%20y%20Can%20USA%2008-18-2026.xlsm&"
    "action=edit&mobileredirect=true"
)
TARGET_SHEET = os.environ.get("WHOLESALE_PDF_TARGET_SHEET", "PO Nuevo")
PRINT_AREA = os.environ.get("WHOLESALE_PDF_PRINT_AREA", "A1:I40")
PDF_FILENAME = os.environ.get("WHOLESALE_PDF_FILENAME", "generar-orden.pdf")
BUTTON_NAME = os.environ.get("WHOLESALE_PDF_BUTTON_NAME", "Ejecutar Generar Orden")
ARTIFACT_DIR = Path(
    os.environ.get(
        "WHOLESALE_ORDER_ARTIFACT_DIR", "artifacts/wholesale_generar_orden"
    )
)


def request_graph(
    method: str,
    url: str,
    headers: dict[str, str],
    *,
    data: bytes | None = None,
    timeout: int = 180,
) -> requests.Response:
    retryable = {409, 423, 429, 500, 502, 503, 504}
    for attempt in range(1, 6):
        response = requests.request(
            method,
            url,
            headers=headers,
            data=data,
            timeout=timeout,
            allow_redirects=True,
        )
        if response.status_code in retryable and attempt < 5:
            retry_after = int(response.headers.get("Retry-After", "0") or 0)
            time.sleep(retry_after or min(attempt * 5, 30))
            continue
        if not response.ok:
            try:
                detail = response.json().get("error", {}).get("message", "")
            except ValueError:
                detail = response.text[:500]
            raise RuntimeError(f"Microsoft Graph respondió {response.status_code}: {detail}")
        return response
    raise RuntimeError("Microsoft Graph no respondió después de los reintentos.")


def download_workbook(token: str, item: dict, destination: Path) -> None:
    drive_id = item["parentReference"]["driveId"]
    response = request_graph(
        "GET",
        f"{GRAPH_URL}/drives/{drive_id}/items/{item['id']}/content",
        graph_headers(token),
        timeout=180,
    )
    destination.write_bytes(response.content)


def _replace_sheet_visibility(workbook_xml: str) -> tuple[str, int, int]:
    sheet_pattern = re.compile(r"<sheet\b[^>]*?(?:/>|></sheet>)")
    sheet_names = re.findall(r'<sheet\b[^>]*\bname="([^"]+)"', workbook_xml)
    if TARGET_SHEET not in sheet_names:
        raise RuntimeError(f"El libro no contiene la hoja {TARGET_SHEET!r}.")
    target_index = sheet_names.index(TARGET_SHEET)

    def rewrite(match: re.Match[str]) -> str:
        tag = match.group(0)
        name_match = re.search(r'\bname="([^"]+)"', tag)
        name = name_match.group(1) if name_match else ""
        tag = re.sub(r'\sstate="[^"]*"', "", tag)
        if name != TARGET_SHEET:
            if tag.rstrip().endswith("/>"):
                tag = tag.rstrip()[:-2].rstrip() + ' state="hidden" />'
            else:
                tag = tag.replace(">", ' state="hidden">', 1)
        return tag

    updated = sheet_pattern.sub(rewrite, workbook_xml)
    active = re.search(r"<workbookView\b[^>]*", updated)
    if not active:
        raise RuntimeError("El libro no contiene workbookView para activar PO Nuevo.")
    active_tag = active.group(0)
    if re.search(r'\bactiveTab="[^"]*"', active_tag):
        updated = updated[: active.start()] + re.sub(
            r'\bactiveTab="[^"]*"',
            f'activeTab="{target_index}"',
            active_tag,
            count=1,
        ) + updated[active.end() :]
    else:
        replacement = active_tag.rstrip() + f' activeTab="{target_index}"'
        updated = updated[: active.start()] + replacement + updated[active.end() :]
    return updated, target_index, len(sheet_names)


def _absolute_print_cell(cell: str) -> str:
    match = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", cell.strip())
    if not match:
        raise RuntimeError(
            f"Área de impresión inválida: {PRINT_AREA!r}. Usa, por ejemplo, A1:I40."
        )
    return f"${match.group(1).upper()}${match.group(2)}"


def _print_area_reference() -> str:
    parts = PRINT_AREA.split(":", 1)
    if len(parts) != 2:
        raise RuntimeError(
            f"Área de impresión inválida: {PRINT_AREA!r}. Usa, por ejemplo, A1:I40."
        )
    escaped_sheet = TARGET_SHEET.replace("'", "''")
    return (
        f"'{escaped_sheet}'!{_absolute_print_cell(parts[0])}:"
        f"{_absolute_print_cell(parts[1])}"
    )


def _replace_print_area(workbook_xml: str, target_index: int) -> str:
    print_area = _print_area_reference()
    updated_names: set[str] = set()
    pattern = re.compile(r"(<definedName\b(?P<attrs>[^>]*)>).*?(</definedName>)")

    def replace_area(match: re.Match[str]) -> str:
        attrs = match.group("attrs")
        name_match = re.search(r'\bname="([^"]+)"', attrs)
        sheet_match = re.search(r'\blocalSheetId="([^"]+)"', attrs)
        name = name_match.group(1) if name_match else ""
        local_sheet_id = sheet_match.group(1) if sheet_match else ""
        if name in {"_xlnm.Print_Area", "Print_Area"} and local_sheet_id == str(target_index):
            updated_names.add(name)
            return f"{match.group(1)}{print_area}{match.group(3)}"
        return match.group(0)

    workbook_xml = pattern.sub(replace_area, workbook_xml)
    if "_xlnm.Print_Area" not in updated_names:
        raise RuntimeError("No se encontró el área de impresión _xlnm.Print_Area de PO Nuevo.")
    return workbook_xml


def _target_worksheet_path(workbook_xml: str, relationships_xml: str) -> str:
    sheet_match = re.search(
        rf'<sheet\b[^>]*\bname="{re.escape(TARGET_SHEET)}"[^>]*\br:id="([^"]+)"',
        workbook_xml,
    )
    if not sheet_match:
        raise RuntimeError(f"No se pudo resolver el XML de la hoja {TARGET_SHEET!r}.")
    relationship_id = sheet_match.group(1)
    relationship = re.search(
        rf'<(?:[A-Za-z0-9_]+:)?Relationship\b[^>]*\bId="{re.escape(relationship_id)}"[^>]*\bTarget="([^"]+)"',
        relationships_xml,
    )
    if not relationship:
        raise RuntimeError(f"No se encontró la relación XML de {TARGET_SHEET!r}.")
    target = relationship.group(1).lstrip("/")
    return target if target.startswith("xl/") else f"xl/{target}"


def _replace_page_setup(sheet_xml: str) -> str:
    pattern = re.compile(r"<pageSetup\b[^>]*/>")
    match = pattern.search(sheet_xml)
    if not match:
        raise RuntimeError(f"La hoja {TARGET_SHEET!r} no contiene configuración de impresión.")
    tag = match.group(0)
    tag = re.sub(r'\sorientation="[^"]*"', '', tag)
    tag = re.sub(r'\spaperSize="[^"]*"', '', tag)
    tag = tag.rstrip()[:-2].rstrip() + ' orientation="portrait" paperSize="1" />'
    return sheet_xml[: match.start()] + tag + sheet_xml[match.end() :]


def prepare_pdf_copy(source: Path, destination: Path) -> dict[str, object]:
    """Mirror the Office Script by changing only the XLSM XML settings.

    Using a ZIP-level patch is intentional.  Loading/saving with a general
    workbook library can discard Excel extensions and cached formula values;
    the temporary PDF copy must retain the original workbook package exactly.
    """
    with ZipFile(source, "r") as source_zip:
        workbook_xml = source_zip.read("xl/workbook.xml").decode("utf-8")
        relationships_xml = source_zip.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        workbook_xml, target_index, sheet_count = _replace_sheet_visibility(workbook_xml)
        workbook_xml = _replace_print_area(workbook_xml, target_index)
        target_path = _target_worksheet_path(workbook_xml, relationships_xml)
        target_sheet_xml = _replace_page_setup(
            source_zip.read(target_path).decode("utf-8")
        )

        with ZipFile(destination, "w") as destination_zip:
            for info in source_zip.infolist():
                content = source_zip.read(info.filename)
                if info.filename == "xl/workbook.xml":
                    content = workbook_xml.encode("utf-8")
                elif info.filename == target_path:
                    content = target_sheet_xml.encode("utf-8")
                destination_zip.writestr(info, content)

    with ZipFile(source, "r") as original_zip, ZipFile(destination, "r") as copy_zip:
        if original_zip.read("xl/vbaProject.bin") != copy_zip.read("xl/vbaProject.bin"):
            raise RuntimeError("La copia temporal no conservó el proyecto VBA.")

    return {
        "sheet": TARGET_SHEET,
        "print_area": PRINT_AREA,
        "orientation": "portrait",
        "paper": "letter",
        "sheet_count": sheet_count,
        "package_mode": "XML patch; formulas, macros y objetos conservados",
    }


def upload_temporary_workbook(token: str, item: dict, workbook_path: Path) -> dict:
    drive_id = item["parentReference"]["driveId"]
    parent_id = item["parentReference"]["id"]
    filename = f".codex-generar-orden-{uuid4().hex}.xlsm"
    upload_url = (
        f"{GRAPH_URL}/drives/{drive_id}/items/{parent_id}:/{quote(filename, safe='')}:/content"
    )
    response = request_graph(
        "PUT",
        upload_url,
        {**graph_headers(token), "Content-Type": "application/octet-stream"},
        data=workbook_path.read_bytes(),
        timeout=180,
    )
    temporary_item = response.json()
    temporary_item["_temporary_name"] = filename
    return temporary_item


def convert_to_pdf(token: str, item: dict, destination: Path) -> None:
    drive_id = item["parentReference"]["driveId"]
    response = request_graph(
        "GET",
        f"{GRAPH_URL}/drives/{drive_id}/items/{item['id']}/content?format=pdf",
        graph_headers(token),
        timeout=240,
    )
    if not response.content.startswith(b"%PDF"):
        raise RuntimeError("Graph no devolvió un PDF válido al convertir el libro.")
    destination.write_bytes(response.content)


def delete_temporary_workbook(token: str, item: dict) -> None:
    drive_id = item["parentReference"]["driveId"]
    request_graph(
        "DELETE",
        f"{GRAPH_URL}/drives/{drive_id}/items/{item['id']}",
        graph_headers(token),
        timeout=60,
    )


def main() -> None:
    ARTIFACT_DIR.mkdir(parents=True, exist_ok=True)
    started = time.time()
    summary: dict[str, object] = {
        "button": BUTTON_NAME,
        "engine": "Microsoft Graph",
        "workbook": os.environ.get("SHAREPOINT_WORKBOOK_URL", DEFAULT_WORKBOOK_URL),
        "status": "error",
    }
    temporary_item: dict | None = None

    try:
        token = graph_token()
        item = resolve_sharepoint_item_by_url(
            token, os.environ.get("SHAREPOINT_WORKBOOK_URL", DEFAULT_WORKBOOK_URL)
        )
        with tempfile.TemporaryDirectory(prefix="wholesale-generar-orden-") as temp_dir:
            source = Path(temp_dir) / "source.xlsm"
            prepared = Path(temp_dir) / "prepared.xlsm"
            download_workbook(token, item, source)
            summary["preparation"] = prepare_pdf_copy(source, prepared)

            temporary_item = upload_temporary_workbook(token, item, prepared)
            summary["temporary_sharepoint_item"] = temporary_item.get("name")
            pdf_path = ARTIFACT_DIR / PDF_FILENAME
            convert_to_pdf(token, temporary_item, pdf_path)
            summary.update(
                {
                    "status": "completed",
                    "pdf": pdf_path.name,
                    "pdf_bytes": pdf_path.stat().st_size,
                    "elapsed_seconds": round(time.time() - started, 1),
                }
            )
            print(f"PDF generado con Graph: {pdf_path}", flush=True)
    except Exception as error:
        summary.update(
            {"error": str(error), "elapsed_seconds": round(time.time() - started, 1)}
        )
        raise
    finally:
        if temporary_item:
            try:
                delete_temporary_workbook(token, temporary_item)
                summary["temporary_sharepoint_item_deleted"] = True
            except Exception as cleanup_error:
                summary["temporary_sharepoint_item_deleted"] = False
                summary["cleanup_error"] = str(cleanup_error)
        (ARTIFACT_DIR / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )


if __name__ == "__main__":
    main()

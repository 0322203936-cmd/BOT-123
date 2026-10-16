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
import shutil
import subprocess
import tempfile
import time
import posixpath
from io import BytesIO
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import quote
from uuid import uuid4
from zipfile import ZipFile

import requests
from openpyxl import load_workbook

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
            if not detail:
                diagnostics = [
                    f"url={response.url}",
                    f"x-errorcode={response.headers.get('x-errorcode', '')}",
                    f"x-correlationid={response.headers.get('x-correlationid', '')}",
                ]
                detail = "; ".join(
                    value for value in diagnostics if not value.endswith("=")
                )
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
        rf'<(?:[A-Za-z0-9_]+:)?Relationship\b(?=[^>]*\bId="{re.escape(relationship_id)}")'
        r'(?=[^>]*\bTarget="([^"]+)")[^>]*/>',
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


def prepare_pdf_xlsx_copy(source: Path, destination: Path) -> None:
    """Create a macro-free temporary copy for Office's PDF converter.

    The source remains an XLSM and is never overwritten. This fallback is only
    used when the Office conversion service cannot open the macro-enabled
    temporary copy; removing VBA avoids conversion failures caused by a
    workbook's macro package while preserving all worksheet data and layout.
    """
    vba_relationship = re.compile(
        rb'<Relationship\b(?=[^>]*\bType="[^"]*/vbaProject(?:"|/))[^>]*/>'
    )
    vba_content_type = re.compile(
        rb'<Override\b(?=[^>]*\bPartName="/xl/vbaProject\.bin")[^>]*/>'
    )
    with ZipFile(source, "r") as source_zip, ZipFile(
        destination, "w"
    ) as destination_zip:
        for info in source_zip.infolist():
            if info.filename == "xl/vbaProject.bin":
                continue
            content = source_zip.read(info.filename)
            if info.filename == "[Content_Types].xml":
                content = content.replace(
                    b"application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                    b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                )
                content = vba_content_type.sub(b"", content)
            elif info.filename == "xl/_rels/workbook.xml.rels":
                content = vba_relationship.sub(b"", content)
            destination_zip.writestr(info, content)


def _read_order_metadata(source: Path) -> dict[str, str]:
    """Read e-mail fields without saving or modifying the source workbook."""
    values_workbook = load_workbook(
        source,
        data_only=True,
        keep_vba=False,
        keep_links=False,
    )
    try:
        sheet_name = "PO Nuevo" if "PO Nuevo" in values_workbook.sheetnames else TARGET_SHEET
        if sheet_name not in values_workbook.sheetnames:
            raise RuntimeError(f"El libro no contiene la hoja {TARGET_SHEET!r}.")
        order_values = values_workbook[sheet_name]
        customer = str(order_values["F11"].value or "Cliente").strip()
        load_date_value = order_values["B19"].value
        if not isinstance(load_date_value, (date, datetime)):
            tender_date = order_values["F19"].value
            if isinstance(tender_date, (date, datetime)):
                load_date_value = tender_date - timedelta(days=1)
        load_date = (
            load_date_value.strftime("%m/%d/%Y")
            if isinstance(load_date_value, (date, datetime))
            else str(load_date_value or "Sin fecha").strip()
        )
        return {"customer": customer or "Cliente", "load_date": load_date or "Sin fecha"}
    finally:
        values_workbook.close()


def _relationship_part_path(part_path: str, relationship_target: str) -> str:
    """Resolve an XLSX relationship target relative to its owning part."""
    relationship_target = relationship_target.lstrip("/")
    if relationship_target.startswith("xl/"):
        return relationship_target
    return posixpath.normpath(
        posixpath.join(posixpath.dirname(part_path), relationship_target)
    )


def _worksheet_relationships_path(worksheet_path: str) -> str:
    return (
        f"{posixpath.dirname(worksheet_path)}/_rels/"
        f"{posixpath.basename(worksheet_path)}.rels"
    )


def _drawing_relationship(xml: str) -> tuple[str, str] | None:
    match = re.search(
        r'<(?:[A-Za-z0-9_]+:)?Relationship\b(?=[^>]*\bType="[^"]*/drawing")'
        r'(?=[^>]*\bId="([^"]+)")(?=[^>]*\bTarget="([^"]+)")[^>]*/>',
        xml,
    )
    return (match.group(1), match.group(2)) if match else None


def _preserve_target_drawing(source: Path, destination: Path) -> None:
    """Restore the original drawing XML after ``openpyxl`` rewrites it.

    The original PO sheet contains native drawing markup for the Pacifica logo.
    ``openpyxl`` keeps the bitmap but rewrites that markup; Microsoft Graph's
    PDF renderer then omits the image.  We keep the safe XLSX package written
    by openpyxl and replace only its already-supported drawing part.
    """
    with ZipFile(source, "r") as source_zip, ZipFile(destination, "r") as clean_zip:
        source_workbook = source_zip.read("xl/workbook.xml").decode("utf-8")
        source_relationships = source_zip.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        clean_workbook = clean_zip.read("xl/workbook.xml").decode("utf-8")
        clean_relationships = clean_zip.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        source_sheet_path = _target_worksheet_path(source_workbook, source_relationships)
        clean_sheet_path = _target_worksheet_path(clean_workbook, clean_relationships)
        source_sheet_rels_path = _worksheet_relationships_path(source_sheet_path)
        clean_sheet_rels_path = _worksheet_relationships_path(clean_sheet_path)
        if source_sheet_rels_path not in source_zip.namelist() or clean_sheet_rels_path not in clean_zip.namelist():
            return
        source_drawing = _drawing_relationship(
            source_zip.read(source_sheet_rels_path).decode("utf-8")
        )
        clean_drawing = _drawing_relationship(
            clean_zip.read(clean_sheet_rels_path).decode("utf-8")
        )
        if not source_drawing or not clean_drawing:
            return
        _, source_drawing_target = source_drawing
        _, clean_drawing_target = clean_drawing
        source_drawing_path = _relationship_part_path(source_sheet_path, source_drawing_target)
        clean_drawing_path = _relationship_part_path(clean_sheet_path, clean_drawing_target)
        source_drawing_rels_path = _worksheet_relationships_path(source_drawing_path)
        clean_drawing_rels_path = _worksheet_relationships_path(clean_drawing_path)
        replacements = {
            clean_drawing_path: source_zip.read(source_drawing_path),
            clean_drawing_rels_path: source_zip.read(source_drawing_rels_path),
        }
        source_drawing_rels = source_zip.read(source_drawing_rels_path).decode("utf-8")
        for target in re.findall(r'\bTarget="([^"]+)"', source_drawing_rels):
            source_part = _relationship_part_path(source_drawing_path, target)
            clean_part = _relationship_part_path(clean_drawing_path, target)
            if source_part in source_zip.namelist():
                replacements[clean_part] = source_zip.read(source_part)

        temporary = destination.with_suffix(".drawing-tmp.xlsx")
        try:
            with ZipFile(temporary, "w") as output_zip:
                for info in clean_zip.infolist():
                    output_zip.writestr(info, replacements.get(info.filename, clean_zip.read(info.filename)))
            # Windows does not allow replacing an archive that is still open.
            clean_zip.close()
            temporary.replace(destination)
        finally:
            temporary.unlink(missing_ok=True)


def _extract_pacifica_logo(workbook_path: Path, destination: Path) -> bool:
    """Extract the Pacifica image embedded in the source Excel workbook."""
    with ZipFile(workbook_path, "r") as workbook_zip:
        image_parts = [
            name
            for name in workbook_zip.namelist()
            if name.startswith("xl/media/") and name.lower().endswith((".png", ".jpg", ".jpeg"))
        ]
        if not image_parts:
            return False
        destination.write_bytes(workbook_zip.read(image_parts[0]))
    return True


def add_pacifica_logo_to_pdf(pdf_path: Path, logo_path: Path) -> None:
    """Draw the logo directly on the first page after Graph conversion.

    Graph converts the worksheet reliably but does not render its embedded
    image.  Adding the same image to the resulting PDF is deterministic and
    keeps the normal Graph conversion fast.
    """
    from pypdf import PdfReader, PdfWriter
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas

    reader = PdfReader(str(pdf_path))
    if not reader.pages:
        raise RuntimeError("El PDF generado no contiene páginas para agregar el logo.")
    first_page = reader.pages[0]
    width = float(first_page.mediabox.width)
    height = float(first_page.mediabox.height)
    overlay_stream = BytesIO()
    overlay = canvas.Canvas(overlay_stream, pagesize=(width, height))
    # Header, directly above the company title.  The fixed position follows
    # the Letter portrait layout used by PO Nuevo and avoids the order table.
    overlay.drawImage(
        ImageReader(str(logo_path)),
        56,
        height - 82,
        width=88,
        height=69,
        preserveAspectRatio=True,
        anchor="sw",
        mask="auto",
    )
    overlay.save()
    overlay_stream.seek(0)
    logo_page = PdfReader(overlay_stream).pages[0]
    first_page.merge_page(logo_page, over=True)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    temporary = pdf_path.with_suffix(".with-logo.pdf")
    try:
        with temporary.open("wb") as output:
            writer.write(output)
        temporary.replace(pdf_path)
    finally:
        temporary.unlink(missing_ok=True)


def prepare_clean_xlsx_copy(source: Path, destination: Path) -> dict[str, str]:
    """Create a Graph-compatible copy while keeping the original logo."""
    metadata = _read_order_metadata(source)
    workbook = load_workbook(
        source,
        data_only=False,
        keep_vba=False,
        keep_links=False,
    )
    values_workbook = load_workbook(
        source,
        data_only=True,
        keep_vba=False,
        keep_links=False,
    )
    try:
        if TARGET_SHEET not in workbook.sheetnames or TARGET_SHEET not in values_workbook.sheetnames:
            raise RuntimeError(f"El libro no contiene la hoja {TARGET_SHEET!r}.")
        values_sheet = values_workbook[TARGET_SHEET]
        for candidate in workbook.worksheets:
            if candidate.title != TARGET_SHEET:
                candidate.sheet_state = "hidden"
        sheet = workbook[TARGET_SHEET]
        for row in sheet.iter_rows():
            for cell in row:
                if cell.data_type == "f":
                    cached_value = values_sheet[cell.coordinate].value
                    if cached_value is not None:
                        cell.value = cached_value
        sheet.print_area = PRINT_AREA
        sheet.page_setup.orientation = "portrait"
        sheet.page_setup.paperSize = sheet.PAPERSIZE_LETTER
        workbook.active = workbook.index(sheet)
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.save(destination)
    finally:
        workbook.close()
        values_workbook.close()
    _preserve_target_drawing(source, destination)
    return metadata


def upload_temporary_workbook(
    token: str, item: dict, workbook_path: Path, *, extension: str = ".xlsm"
) -> dict:
    drive_id = item["parentReference"]["driveId"]
    parent_id = item["parentReference"]["id"]
    # Keep the temporary item out of the normal workbook view, but do not use
    # a dot-prefixed filename. Graph's Excel conversion service can reject
    # hidden drive items with HTTP 406 even though the upload itself succeeds.
    filename = f"codex-generar-orden-{uuid4().hex}{extension}"
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
    url = f"{GRAPH_URL}/drives/{drive_id}/items/{item['id']}/content?format=pdf"
    # Ask Graph explicitly for the PDF representation. Without Accept, the
    # conversion endpoint can answer 406 with an empty body.
    pdf_headers = {
        **graph_headers(token),
        "Accept": "application/pdf",
        "Cache-Control": "no-cache",
    }

    last_error: RuntimeError | None = None
    for attempt in range(1, 5):
        try:
            response = request_graph(
                "GET",
                url,
                pdf_headers,
                timeout=240,
            )
            if not response.content.startswith(b"%PDF"):
                raise RuntimeError("Graph no devolvió un PDF válido al convertir el libro.")
            destination.write_bytes(response.content)
            return
        except RuntimeError as error:
            last_error = error
            if "respondió 406" not in str(error) or attempt == 4:
                raise
            delay = min(attempt * 5, 20)
            print(
                f"Graph aún no ofrece el PDF (intento {attempt}/4); "
                f"reintentando en {delay}s...",
                flush=True,
            )
            time.sleep(delay)

    if last_error is not None:
        raise last_error


def convert_with_libreoffice(source: Path, destination: Path) -> None:
    """Convert the prepared local copy when Graph's Office service rejects it."""
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    if not executable:
        install = subprocess.run(
            ["sudo", "apt-get", "update"],
            capture_output=True,
            text=True,
            timeout=300,
            check=False,
        )
        if install.returncode == 0:
            install = subprocess.run(
                [
                    "sudo",
                    "apt-get",
                    "install",
                    "-y",
                    "--no-install-recommends",
                    "libreoffice-calc",
                ],
                capture_output=True,
                text=True,
                timeout=600,
                check=False,
            )
        executable = shutil.which("libreoffice") or shutil.which("soffice")
        if install.returncode != 0 or not executable:
            detail = (install.stderr or install.stdout).strip()[-500:]
            raise RuntimeError(
                "No se pudo preparar LibreOffice para el respaldo local de PDF: "
                f"{detail}"
            )

    destination.parent.mkdir(parents=True, exist_ok=True)
    generated = destination.parent / f"{source.stem}.pdf"
    if generated.exists():
        generated.unlink()
    result = subprocess.run(
        [
            executable,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(destination.parent),
            str(source),
        ],
        capture_output=True,
        text=True,
        timeout=240,
        check=False,
    )
    if result.returncode != 0 or not generated.exists():
        detail = (result.stderr or result.stdout).strip()[-500:]
        raise RuntimeError(f"LibreOffice no pudo generar el PDF: {detail}")
    generated.replace(destination)
    if not destination.read_bytes().startswith(b"%PDF"):
        raise RuntimeError("LibreOffice no devolvió un PDF válido.")


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
    temporary_items: list[dict] = []

    try:
        token = graph_token()
        item = resolve_sharepoint_item_by_url(
            token, os.environ.get("SHAREPOINT_WORKBOOK_URL", DEFAULT_WORKBOOK_URL)
        )
        with tempfile.TemporaryDirectory(prefix="wholesale-generar-orden-") as temp_dir:
            source = Path(temp_dir) / "source.xlsm"
            download_workbook(token, item, source)
            logo_path = Path(temp_dir) / "pacifica-logo.png"
            has_logo = _extract_pacifica_logo(source, logo_path)
            pdf_path = ARTIFACT_DIR / PDF_FILENAME
            clean_xlsx = Path(temp_dir) / "prepared-for-pdf.xlsx"
            order_metadata = prepare_clean_xlsx_copy(source, clean_xlsx)
            document_kind = "Orden" if TARGET_SHEET == "PO Nuevo" else "Passing Sheet"
            summary["order_email"] = {
                **order_metadata,
                "document_kind": document_kind,
            }
            summary["preparation"] = {
                "sheet": TARGET_SHEET,
                "print_area": PRINT_AREA,
                "orientation": "portrait",
                "paper": "letter",
                "package_mode": "macro-free copy; lookup sheets hidden",
            }
            clean_item = upload_temporary_workbook(
                token, item, clean_xlsx, extension=".xlsx"
            )
            temporary_items.append(clean_item)
            summary["temporary_sharepoint_item"] = clean_item.get("name")
            try:
                # This clean workbook is the proven successful conversion path.
                # Starting here avoids two 35-second 406 retries from the
                # macro-enabled and ZIP-patched copies used previously.
                convert_to_pdf(token, clean_item, pdf_path)
                summary["pdf_fallback"] = "fast clean PO Nuevo XLSX"
            except RuntimeError as conversion_error:
                if "respondió 406" not in str(conversion_error):
                    raise
                print(
                    "Graph no pudo convertir la copia rápida; usando LibreOffice "
                    "local como respaldo...",
                    flush=True,
                )
                convert_with_libreoffice(clean_xlsx, pdf_path)
                summary["pdf_fallback"] = "LibreOffice from fast clean PO Nuevo XLSX"
            if TARGET_SHEET == "PO Nuevo" and has_logo:
                add_pacifica_logo_to_pdf(pdf_path, logo_path)
                summary["logo"] = "Pacifica logo added directly to final PDF"
            elif TARGET_SHEET == "PO Nuevo":
                summary["logo"] = "Pacifica logo not found in source workbook"
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
        cleanup_errors: list[str] = []
        for temporary_item in reversed(temporary_items):
            try:
                delete_temporary_workbook(token, temporary_item)
            except Exception as cleanup_error:
                cleanup_errors.append(str(cleanup_error))
        summary["temporary_sharepoint_items_deleted"] = len(temporary_items) - len(
            cleanup_errors
        )
        if cleanup_errors:
            summary["temporary_sharepoint_item_deleted"] = False
            summary["cleanup_error"] = " | ".join(cleanup_errors)
        elif temporary_items:
            summary["temporary_sharepoint_item_deleted"] = True
        (ARTIFACT_DIR / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )


if __name__ == "__main__":
    main()

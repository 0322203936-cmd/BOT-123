"""Create the POSCO import workbook for a Wholesale order.

This module intentionally does not alter the existing Facturas Komet bot.  It
only reuses its approved POSCO template and column contract, while reading the
current ``PO Nuevo`` and ``Productos`` sheets from the Wholesale workbook.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path
import re

from openpyxl import load_workbook

from posco_order_transform import (
    TransformationError,
    carrier_abbreviation,
    clean,
    required_positive_int,
    secret_workbook,
    write_order_workbook,
)


PRODUCT_HEADERS = (
    "Distribuidor",
    "Cust Name",
    "CATEGORIA",
    "FLOR",
    "Flor Color",
    "Descripcion",
    "Ramos/Cajas",
    "txr_orden",
    "Caja",
)


def normalize_description(value: object) -> str:
    """Compare product descriptions without cosmetic casing or spacing noise."""
    text = clean(value).casefold()
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def is_boom_pending(value: object) -> bool:
    """The PO Nuevo status column is the sole source of BOOM pending alerts."""
    status = clean(value).casefold()
    return status in {"pendiente de boom", "pendiente de bom"}


def as_date(value: object, label: str) -> datetime:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    raise TransformationError(f"{label} está vacío o no es una fecha válida en PO Nuevo.")


def required_wholesale_quantity(value: object, label: str) -> int:
    """Accept the numeric labels shown by PO Nuevo, such as ``6ST`` or ``10BU``."""
    try:
        return required_positive_int(value, label)
    except TransformationError as original_error:
        text = clean(value)
        match = re.fullmatch(
            r"(\d+)\s*(?:st(?:em(?:s)?)?|bu(?:nch(?:es)?)?|pk|box)?", text,
            re.IGNORECASE,
        )
        if not match:
            raise original_error
        return required_positive_int(match.group(1), label)


def build_boom_recommendation(
    flower: object,
    color: object,
    stems: object,
    bunches: object,
    box: object,
) -> str:
    """Return a readable BOOM proposal from the manually supplied PO values."""
    parts = [clean(flower)]
    if clean(color):
        parts.append(clean(color))
    if clean(stems):
        parts.append(f"STEM/{clean(stems)}")
    if clean(bunches):
        parts.append(f"BU/{clean(bunches)}")
    if clean(box):
        parts.append(clean(box))
    return " ".join(part for part in parts if part) or "Producto Wholesale"


def _header_indexes(sheet) -> dict[str, int]:
    headings = [clean(sheet.cell(1, column).value) for column in range(1, 10)]
    if tuple(headings) != PRODUCT_HEADERS:
        raise TransformationError(
            "La hoja Productos no tiene las columnas A:I esperadas para Wholesale."
        )
    return {heading.casefold(): index for index, heading in enumerate(headings)}


def load_products_lookup(sheet) -> dict[str, list[dict[str, str]]]:
    indexes = _header_indexes(sheet)
    lookup: dict[str, list[dict[str, str]]] = {}
    for values in sheet.iter_rows(min_row=2, max_col=9, values_only=True):
        entry = {
            header.casefold(): clean(values[index])
            for header, index in indexes.items()
        }
        description = entry["descripcion"]
        if not description:
            continue
        lookup.setdefault(normalize_description(description), []).append(entry)
    if not lookup:
        raise TransformationError("La hoja Productos no contiene productos homologados.")
    return lookup


def _select_product(
    lookup: dict[str, list[dict[str, str]]],
    description: str,
    color: str,
    box: str,
) -> dict[str, str] | None:
    """Resolve a known product using its description, color and box.

    POSCO descriptions can contain harmless spelling differences (for example
    ``ASSORTED`` versus the historic ``ASSORED`` recorded in Productos).  A
    close match is allowed only after the supplied color and box reduce it to
    one product, so similar products never become a false homologation.
    """
    candidates = lookup.get(normalize_description(description), [])
    if not candidates:
        normalized = normalize_description(description)
        close_keys = [
            key
            for key in lookup
            if SequenceMatcher(None, normalized, key).ratio() >= 0.92
        ]
        candidates = [candidate for key in close_keys for candidate in lookup[key]]
    if not candidates:
        return None
    exact_box = [candidate for candidate in candidates if clean(candidate["caja"]).casefold() == box.casefold()]
    selected = exact_box or candidates
    if color:
        exact_color = [
            candidate
            for candidate in selected
            if clean(candidate["flor color"]).casefold() == color.casefold()
        ]
        selected = exact_color or selected
    # A description must map to one unique product.  A similarly named item
    # with a different box is ambiguous and must wait for its BOOM mapping.
    return selected[0] if len(selected) == 1 else None


def generate_posco_from_po_new(source: Path, destination: Path) -> dict[str, object]:
    """Build a POSCO-ready XLSX plus a list of products pending BOOM creation."""
    workbook = load_workbook(source, data_only=True, keep_links=False)
    try:
        if "PO Nuevo" not in workbook.sheetnames or "Productos" not in workbook.sheetnames:
            raise TransformationError("El libro Wholesale debe incluir las hojas PO Nuevo y Productos.")
        order_sheet = workbook["PO Nuevo"]
        products = load_products_lookup(workbook["Productos"])

        po_number = clean(order_sheet["A9"].value)
        customer = clean(order_sheet["F11"].value)
        truck_line = clean(order_sheet["F15"].value)
        load_date = as_date(order_sheet["B19"].value, "Load Date")
        if not po_number or not customer or not truck_line:
            raise TransformationError("PO Nuevo requiere PO, Customer Ship y Truck Line para crear el formato POSCO.")
        no_value = f"{customer} - {carrier_abbreviation(truck_line)} {po_number}"

        rows: list[list[object]] = []
        boom_pending: list[dict[str, str]] = []
        for row_number in range(22, 41):
            description_input = clean(order_sheet.cell(row_number, 2).value)
            if not description_input:
                continue
            color_input = clean(order_sheet.cell(row_number, 3).value)
            stems_input = order_sheet.cell(row_number, 4).value
            bunches_input = order_sheet.cell(row_number, 5).value
            box_input = clean(order_sheet.cell(row_number, 6).value)
            stems = required_wholesale_quantity(stems_input, f"STEM/BCH de la fila {row_number}")
            bunches = required_wholesale_quantity(bunches_input, f"BU/BOX de la fila {row_number}")
            product = _select_product(products, description_input, color_input, box_input)
            pending_boom = is_boom_pending(order_sheet.cell(row_number, 10).value)

            if product:
                description = product["descripcion"]
                category = product["categoria"] or "CB / BULK"
                flower = product["flor"] or description_input
                color = product["flor color"] or color_input
                box = product["caja"] or box_input
            else:
                description = description_input
                category = "CB / BULK"
                flower = description_input
                color = color_input
                box = box_input
            if pending_boom:
                boom_pending.append(
                    {
                        "product": description_input,
                        "recommendation": build_boom_recommendation(
                            description_input, color_input, stems, bunches, box_input
                        ),
                    }
                )

            rows.append(
                [
                    "PACIFICA PRODUCE FARMS",
                    "WHOLESALE",
                    description,
                    no_value,
                    category,
                    flower,
                    color,
                    load_date,
                    1,
                    bunches,
                    stems,
                    load_date - timedelta(days=3),
                    box,
                ]
            )

        if not rows:
            raise TransformationError("PO Nuevo no contiene productos para crear el formato POSCO.")
        write_order_workbook(
            secret_workbook("KOMET_POSCO_TEMPLATE_B64"),
            rows,
            destination,
        )
        return {
            "file": destination.name,
            "rows": len(rows),
            "boom_pending": boom_pending,
            "no": no_value,
        }
    finally:
        workbook.close()

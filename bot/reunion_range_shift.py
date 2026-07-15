import os
from datetime import date, datetime
from typing import Any

from excel_range_sync import graph_request
from sharepoint_sync import GRAPH_URL, graph_headers, graph_token, resolve_sharepoint_item


SHEET_NAME = "Reunion"
DATE_CELL = "D1"
HEADER_ROW = 2
FIRST_DATA_ROW = 3


def env_flag(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def column_letter(number: int) -> str:
    if number < 1:
        raise ValueError("El número de columna debe ser positivo.")
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def normalized(value: Any) -> str:
    return str(value or "").strip().casefold()


def find_cor_columns(headers: list[Any], first_column: int = 1) -> list[int]:
    columns = [
        first_column + index
        for index, value in enumerate(headers)
        if normalized(value) == "cor"
    ]
    if len(columns) < 2:
        raise RuntimeError("Se esperaban al menos dos columnas con encabezado Cor.")
    gaps = [right - left for left, right in zip(columns, columns[1:])]
    if any(gap != 6 for gap in gaps):
        labels = ", ".join(column_letter(column) for column in columns)
        raise RuntimeError(
            f"Las columnas Cor no conservan bloques de seis columnas: {labels}."
        )
    return columns


def find_last_data_row(values: list[list[Any]], start_row: int = FIRST_DATA_ROW) -> int:
    last_row = start_row - 1
    for offset, row in enumerate(values):
        value = row[0] if row else ""
        if normalized(value):
            last_row = start_row + offset
    if last_row < start_row:
        raise RuntimeError("No se encontraron flores en la columna A.")
    return last_row


def next_excel_date(value: Any) -> int | float:
    if isinstance(value, bool):
        raise RuntimeError("D1 no contiene una fecha válida.")
    if isinstance(value, (int, float)):
        return value + 1
    text = str(value or "").strip()
    for pattern in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            parsed = datetime.strptime(text, pattern).date()
            return (parsed - date(1899, 12, 30)).days + 1
        except ValueError:
            continue
    raise RuntimeError(f"D1 no contiene una fecha reconocible: {text or 'vacía'}.")


def range_url(workbook_url: str, address: str) -> str:
    return (
        f"{workbook_url}/worksheets/{SHEET_NAME}/"
        f"range(address='{address}')"
    )


def read_range(
    workbook_url: str, headers: dict[str, str], address: str
) -> dict[str, Any]:
    return graph_request(
        "GET", range_url(workbook_url, address), headers, timeout=60
    ).json()


def vertical_values(payload: dict[str, Any], count: int) -> list[list[Any]]:
    source = payload.get("values", [])
    return [[source[index][0] if index < len(source) and source[index] else ""] for index in range(count)]


def update_reunion(token: str, item: dict, apply_changes: bool = True) -> dict[str, Any]:
    drive_id = item["parentReference"]["driveId"]
    workbook_url = f"{GRAPH_URL}/drives/{drive_id}/items/{item['id']}/workbook"
    headers = {**graph_headers(token), "Content-Type": "application/json"}
    session = graph_request(
        "POST",
        f"{workbook_url}/createSession",
        headers,
        json={"persistChanges": apply_changes},
        timeout=60,
    ).json()
    session_headers = {**headers, "workbook-session-id": session["id"]}

    try:
        used_range = graph_request(
            "GET",
            f"{workbook_url}/worksheets/{SHEET_NAME}/usedRange(valuesOnly=false)",
            session_headers,
            timeout=60,
        ).json()
        first_column = int(used_range.get("columnIndex") or 0) + 1
        last_column = first_column + int(used_range.get("columnCount") or 1) - 1
        first_row = int(used_range.get("rowIndex") or 0) + 1
        last_used_row = first_row + int(used_range.get("rowCount") or 1) - 1

        header_address = (
            f"{column_letter(first_column)}{HEADER_ROW}:"
            f"{column_letter(last_column)}{HEADER_ROW}"
        )
        header_rows = read_range(
            workbook_url, session_headers, header_address
        ).get("values", [[]])
        cor_columns = find_cor_columns(header_rows[0] if header_rows else [], first_column)

        flower_values = read_range(
            workbook_url,
            session_headers,
            f"A{FIRST_DATA_ROW}:A{last_used_row}",
        ).get("values", [])
        last_data_row = find_last_data_row(flower_values)
        row_count = last_data_row - FIRST_DATA_ROW + 1

        date_payload = read_range(workbook_url, session_headers, DATE_CELL)
        date_rows = date_payload.get("values", [[]])
        current_date = date_rows[0][0] if date_rows and date_rows[0] else None
        new_date = next_excel_date(current_date)

        snapshots: dict[int, list[list[Any]]] = {}
        for column in cor_columns:
            label = column_letter(column)
            payload = read_range(
                workbook_url,
                session_headers,
                f"{label}{FIRST_DATA_ROW}:{label}{last_data_row}",
            )
            formulas = payload.get("formulas", [])
            formula_cells = [
                FIRST_DATA_ROW + index
                for index, row in enumerate(formulas)
                if row and isinstance(row[0], str) and row[0].startswith("=")
            ]
            if formula_cells:
                raise RuntimeError(
                    f"La columna {label} contiene fórmulas dentro de las filas de flores."
                )
            snapshots[column] = vertical_values(payload, row_count)

        labels = [column_letter(column) for column in cor_columns]
        result = {
            "columns": labels,
            "first_data_row": FIRST_DATA_ROW,
            "last_data_row": last_data_row,
            "current_date": current_date,
            "new_date": new_date,
            "applied": apply_changes,
        }
        print(
            "REUNION_PLAN_OK "
            f"columnas_cor={','.join(labels)} filas={FIRST_DATA_ROW}:{last_data_row} "
            f"fecha_actual={current_date} fecha_nueva={new_date} aplicar={apply_changes}",
            flush=True,
        )
        if not apply_changes:
            return result

        for target, source in zip(cor_columns, cor_columns[1:]):
            target_label = column_letter(target)
            graph_request(
                "PATCH",
                range_url(
                    workbook_url,
                    f"{target_label}{FIRST_DATA_ROW}:{target_label}{last_data_row}",
                ),
                session_headers,
                json={"values": snapshots[source]},
                timeout=90,
            )

        last_label = column_letter(cor_columns[-1])
        graph_request(
            "PATCH",
            range_url(
                workbook_url,
                f"{last_label}{FIRST_DATA_ROW}:{last_label}{last_data_row}",
            ),
            session_headers,
            json={"values": [[0] for _ in range(row_count)]},
            timeout=90,
        )
        graph_request(
            "PATCH",
            range_url(workbook_url, DATE_CELL),
            session_headers,
            json={"values": [[new_date]]},
            timeout=60,
        )

        verified_date_rows = read_range(
            workbook_url, session_headers, DATE_CELL
        ).get("values", [[]])
        verified_date = (
            verified_date_rows[0][0]
            if verified_date_rows and verified_date_rows[0]
            else None
        )
        verified_last = vertical_values(
            read_range(
                workbook_url,
                session_headers,
                f"{last_label}{FIRST_DATA_ROW}:{last_label}{last_data_row}",
            ),
            row_count,
        )
        if verified_date != new_date:
            raise RuntimeError("La fecha D1 no quedó actualizada correctamente.")
        if any(row[0] != 0 for row in verified_last):
            raise RuntimeError(f"La última columna Cor ({last_label}) no quedó en cero.")

        print(
            "REUNION_SHIFT_OK "
            f"hoja={SHEET_NAME} fecha_D1_incrementada=true "
            f"columnas_cor={','.join(labels)} filas={FIRST_DATA_ROW}:{last_data_row} "
            "calculos_inferiores_intactos=true libro_no_reemplazado=true",
            flush=True,
        )
        return result
    finally:
        try:
            graph_request(
                "POST",
                f"{workbook_url}/closeSession",
                session_headers,
                timeout=30,
            )
        except Exception as exc:
            print(f"Aviso al cerrar la sesión de Reunion: {exc}", flush=True)


def main() -> None:
    token = graph_token()
    item = resolve_sharepoint_item(token)
    update_reunion(
        token,
        item,
        apply_changes=env_flag("REUNION_APPLY_CHANGES", default=True),
    )


if __name__ == "__main__":
    main()

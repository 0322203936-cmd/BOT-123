from typing import Any, Callable
from urllib.parse import quote


APPEND_TABLE = "Append1"
PIVOT_SHEET = "Weeks x FechaProduccion"
PIVOT_TABLE = "PivotTable6"
SOURCE_TABLES = ("TableProy", "TableReq")
CHUNK_SIZE = 200

GraphRequest = Callable[..., Any]


def excel_rows_equal(expected: list[Any], actual: list[Any]) -> bool:
    """Compare Excel rows treating Graph's null and blank as the same cell."""
    if len(expected) != len(actual):
        return False

    def normalized(value: Any) -> Any:
        return "" if value is None else value

    return all(
        normalized(expected_value) == normalized(actual_value)
        for expected_value, actual_value in zip(expected, actual)
    )


def _normalized_rows(payload: dict[str, Any], table_name: str) -> list[list[Any]]:
    values = payload.get("values", [])
    if not values or not isinstance(values[0], list):
        raise RuntimeError(f"La tabla {table_name} no devolvio encabezados.")

    width = len(values[0])
    if width < 1:
        raise RuntimeError(f"La tabla {table_name} no contiene columnas.")

    normalized = []
    for row in values:
        if not isinstance(row, list):
            raise RuntimeError(f"La tabla {table_name} devolvio una fila invalida.")
        normalized.append((row + [""] * width)[:width])
    return normalized


def combine_like_power_query(
    table_proy: list[list[Any]],
    table_req: list[list[Any]],
    append_headers: list[Any],
) -> list[list[Any]]:
    """Reproduce Power Query: Table.Combine({TableProy, TableReq})."""
    if not table_proy or not table_req or not append_headers:
        raise RuntimeError("No hay datos suficientes para reconstruir Append1.")

    target_headers = [str(value) for value in append_headers]
    if len(target_headers) != len(set(target_headers)):
        raise RuntimeError("Append1 contiene encabezados duplicados.")

    combined: list[list[Any]] = []
    for table_name, source in zip(SOURCE_TABLES, (table_proy, table_req)):
        source_headers = [str(value) for value in source[0]]
        if len(source_headers) != len(set(source_headers)):
            raise RuntimeError(f"{table_name} contiene encabezados duplicados.")

        unknown = [header for header in source_headers if header not in target_headers]
        if unknown:
            raise RuntimeError(
                f"{table_name} contiene columnas que Append1 no reconoce: "
                + ", ".join(unknown)
            )

        index_by_header = {
            header: index for index, header in enumerate(source_headers)
        }
        for source_row in source[1:]:
            row = (source_row + [""] * len(source_headers))[: len(source_headers)]
            combined.append(
                [
                    row[index_by_header[header]]
                    if header in index_by_header
                    else None
                    for header in target_headers
                ]
            )
    return combined


def _table_range(
    graph_request: GraphRequest,
    workbook_url: str,
    headers: dict[str, str],
    table_name: str,
) -> dict[str, Any]:
    return graph_request(
        "GET",
        f"{workbook_url}/tables/{table_name}/range",
        headers,
        timeout=120,
    ).json()


def _add_rows(
    graph_request: GraphRequest,
    workbook_url: str,
    headers: dict[str, str],
    count: int,
    width: int,
) -> None:
    remaining = count
    while remaining:
        batch = min(CHUNK_SIZE, remaining)
        graph_request(
            "POST",
            f"{workbook_url}/tables/{APPEND_TABLE}/rows/add",
            headers,
            json={"index": None, "values": [[""] * width for _ in range(batch)]},
            timeout=120,
        )
        remaining -= batch


def _delete_surplus_rows(
    graph_request: GraphRequest,
    workbook_url: str,
    headers: dict[str, str],
    desired_data_rows: int,
    current_data_rows: int,
) -> None:
    if desired_data_rows >= current_data_rows:
        return

    first_surplus_row = desired_data_rows + 2
    last_current_row = current_data_rows + 1
    graph_request(
        "POST",
        (
            f"{workbook_url}/worksheets/{quote(APPEND_TABLE, safe='')}/"
            f"range(address='{first_surplus_row}:{last_current_row}')/delete"
        ),
        headers,
        json={"shift": "Up"},
        timeout=120,
    )


def rebuild_append1(
    graph_request: GraphRequest,
    workbook_url: str,
    headers: dict[str, str],
) -> int:
    proy = _normalized_rows(
        _table_range(graph_request, workbook_url, headers, SOURCE_TABLES[0]),
        SOURCE_TABLES[0],
    )
    req = _normalized_rows(
        _table_range(graph_request, workbook_url, headers, SOURCE_TABLES[1]),
        SOURCE_TABLES[1],
    )
    append = _normalized_rows(
        _table_range(graph_request, workbook_url, headers, APPEND_TABLE),
        APPEND_TABLE,
    )

    append_headers = append[0]
    combined = combine_like_power_query(proy, req, append_headers)
    current_data_rows = len(append) - 1
    desired_data_rows = len(combined)
    width = len(append_headers)

    if desired_data_rows > current_data_rows:
        print(
            f"Ampliando Append1 en {desired_data_rows - current_data_rows} filas...",
            flush=True,
        )
        _add_rows(
            graph_request,
            workbook_url,
            headers,
            desired_data_rows - current_data_rows,
            width,
        )
    elif desired_data_rows < current_data_rows:
        print(
            f"Reduciendo Append1 en {current_data_rows - desired_data_rows} filas...",
            flush=True,
        )
        _delete_surplus_rows(
            graph_request,
            workbook_url,
            headers,
            desired_data_rows,
            current_data_rows,
        )

    print(f"Actualizando Append1 con {desired_data_rows} filas...", flush=True)
    for offset in range(0, desired_data_rows, CHUNK_SIZE):
        rows = combined[offset : offset + CHUNK_SIZE]
        start_row = offset + 2
        end_row = start_row + len(rows) - 1
        graph_request(
            "PATCH",
            (
                f"{workbook_url}/worksheets/{quote(APPEND_TABLE, safe='')}/"
                f"range(address='A{start_row}:AA{end_row}')"
            ),
            headers,
            json={"values": rows},
            timeout=120,
        )

    verification = _normalized_rows(
        _table_range(graph_request, workbook_url, headers, APPEND_TABLE),
        APPEND_TABLE,
    )
    if len(verification) - 1 != desired_data_rows:
        raise RuntimeError(
            "Append1 no quedo con la cantidad esperada de filas "
            f"({len(verification) - 1} != {desired_data_rows})."
        )
    if combined and not excel_rows_equal(combined[-1], verification[-1]):
        raise RuntimeError("La ultima fila de Append1 no coincide con el resultado esperado.")

    print(
        f"APPEND1_UPDATE_OK filas={desired_data_rows} "
        "origen=TableProy+TableReq",
        flush=True,
    )
    return desired_data_rows


def refresh_weeks_pivot(
    graph_request: GraphRequest,
    workbook_url: str,
    headers: dict[str, str],
) -> None:
    sheet = quote(PIVOT_SHEET, safe="")
    pivot = quote(PIVOT_TABLE, safe="")

    pivot_info = graph_request(
        "GET",
        (
            f"{workbook_url}/worksheets/{sheet}/"
            f"pivotTables/{pivot}"
        ),
        headers,
        timeout=60,
    ).json()
    if pivot_info.get("name") != PIVOT_TABLE:
        raise RuntimeError(
            f"No se encontro {PIVOT_TABLE} en la hoja {PIVOT_SHEET}."
        )

    graph_request(
        "POST",
        (
            f"{workbook_url}/worksheets/{sheet}/"
            f"pivotTables/{pivot}/refresh"
        ),
        headers,
        timeout=120,
    )
    print(
        f"PIVOT_REFRESH_OK hoja={PIVOT_SHEET} tabla={PIVOT_TABLE}",
        flush=True,
    )


def refresh_data_proy_outputs(
    graph_request: GraphRequest,
    workbook_url: str,
    headers: dict[str, str],
) -> int:
    rows = rebuild_append1(graph_request, workbook_url, headers)

    graph_request(
        "POST",
        f"{workbook_url}/application/calculate",
        headers,
        json={"calculationType": "Full"},
        timeout=120,
    )
    refresh_weeks_pivot(graph_request, workbook_url, headers)
    return rows

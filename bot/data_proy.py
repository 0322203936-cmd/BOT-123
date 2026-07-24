import os
import re
import sys
from pathlib import Path

from sharepoint_sync import (
    GRAPH_URL,
    graph_headers,
    graph_token,
    resolve_sharepoint_item_by_url,
)
from excel_range_sync import graph_request

REQ_PROY_URL = "https://pacificafarms.sharepoint.com/:x:/r/sites/requerimientovsproyeccion/_layouts/15/Doc.aspx?sourcedoc=%7B277A76AA-508A-47F8-8A4A-F19D46660D65%7D&file=Requerimiento%20vs%20proyeccion%20Test.xlsm&action=default&mobileredirect=true"
PLAN_COSECHA_URL = "https://pacificafarms.sharepoint.com/:x:/r/sites/requerimientovsproyeccion/_layouts/15/Doc.aspx?sourcedoc=%7B0A3464AB-7BD8-400A-A0E6-5BC92E23CE3E%7D&file=Plan%20de%20cosecha%202026%20Test.xlsx&action=default&mobileredirect=true"


def col_letter_to_index(letter):
    idx = 0
    for char in letter.upper():
        idx = idx * 26 + (ord(char) - ord('A') + 1)
    return idx - 1


def make_blocks(row_list):
    blocks, cur = [], []
    for r in sorted(row_list):
        if not cur or r == cur[-1] + 1:
            cur.append(r)
        else:
            blocks.append(cur)
            cur = [r]
    if cur:
        blocks.append(cur)
    return blocks


def patch(workbook_url, sh, addr, values):
    """PATCH a range and immediately GET it back to verify the write landed."""
    graph_request("PATCH", f"{workbook_url}/worksheets/DataProy/range(address='{addr}')",
                  sh, json={"values": values})


def clear_and_write(workbook_url, sh, sr, er, cnt,
                    fcd_vals, tallos_vals, sem_vals):
    """
    For each block of rows:
      1. Overwrite O and S with None (null) to clear numeric cells.
      2. Immediately write the real values in the same session.
    Using None (null JSON) is the only reliable way to blank a
    number-formatted cell via Graph API without triggering a type mismatch.
    """
    # Step A – clear O and S with null so Excel accepts the blank
    patch(workbook_url, sh, f"O{sr}:O{er}", [[None] for _ in range(cnt)])
    patch(workbook_url, sh, f"S{sr}:S{er}", [[None] for _ in range(cnt)])

    # Step B – write the new values (numbers, not strings)
    patch(workbook_url, sh, f"F{sr}:H{er}", fcd_vals)
    patch(workbook_url, sh, f"O{sr}:O{er}", tallos_vals)
    patch(workbook_url, sh, f"S{sr}:S{er}", sem_vals)


def main():
    print("Obteniendo token de Microsoft Graph...")
    token = graph_token()

    print("Resolviendo archivos en SharePoint...")
    item_req  = resolve_sharepoint_item_by_url(token, REQ_PROY_URL)
    item_plan = resolve_sharepoint_item_by_url(token, PLAN_COSECHA_URL)

    # -----------------------------------------------------------------------
    # PASO 1: Leer Plan de Cosecha en vivo via Graph API
    # -----------------------------------------------------------------------
    print("Abriendo sesion en vivo de Plan de Cosecha...")
    plan_drive_id     = item_plan["parentReference"]["driveId"]
    plan_workbook_url = f"{GRAPH_URL}/drives/{plan_drive_id}/items/{item_plan['id']}/workbook"
    headers           = {**graph_headers(token), "Content-Type": "application/json"}

    plan_sess = graph_request(
        "POST", f"{plan_workbook_url}/createSession", headers,
        json={"persistChanges": False}, timeout=60,
    ).json()
    psh = {**headers, "workbook-session-id": plan_sess["id"]}

    try:
        ws_list    = graph_request("GET", f"{plan_workbook_url}/worksheets", psh).json()
        sheetnames = [ws["name"] for ws in ws_list.get("value", [])]

        cosecha_sheets = [s for s in sheetnames if str(s).startswith("P Cosecha ")]
        if not cosecha_sheets:
            print("Error: No se encontraron hojas de 'P Cosecha'")
            sys.exit(1)

        latest_sheet_name = sorted(cosecha_sheets)[-1]
        print(f"Hoja detectada (en vivo): {latest_sheet_name}")

        from urllib.parse import quote
        quoted_sheet = quote(latest_sheet_name)

        print("Obteniendo datos en vivo...")
        range_data  = graph_request(
            "GET",
            f"{plan_workbook_url}/worksheets('{quoted_sheet}')/range(address='A1:X1000')",
            psh
        ).json()
        plan_values = range_data.get("values", [])

        def get_week(col):
            if len(plan_values) <= 3:
                return ""
            if col - 1 >= len(plan_values[3]):
                return ""
            val = plan_values[3][col - 1]
            if val is None or str(val).strip() == "":
                return ""
            nums = re.findall(r'\d+', str(val))
            return int(nums[0]) if nums else str(val).strip()

        weeks = [get_week(6),  get_week(18), get_week(19), get_week(20),
                 get_week(21), get_week(22), get_week(23), get_week(24)]
        weeks = [w for w in weeks if w != ""]   # drop empty slots
        print(f"Semanas a procesar: {weeks}")

        flowers_data = []
        for row_idx in range(4, len(plan_values)):
            row_data = plan_values[row_idx]
            if not row_data:
                continue
            flor_val = row_data[0]
            if flor_val is not None and str(flor_val).strip().lower() == "total":
                break
            if not flor_val:
                continue
            flor_real  = row_data[3] if len(row_data) > 3 else None
            color_real = row_data[4] if len(row_data) > 4 else None
            if not flor_real:
                continue
            flowers_data.append({
                "flor":  flor_real,
                "color": color_real,
                "qtys": [
                    row_data[5]  if len(row_data) > 5  and row_data[5]  is not None else 0,
                    row_data[17] if len(row_data) > 17 and row_data[17] is not None else 0,
                    row_data[18] if len(row_data) > 18 and row_data[18] is not None else 0,
                    row_data[19] if len(row_data) > 19 and row_data[19] is not None else 0,
                    row_data[20] if len(row_data) > 20 and row_data[20] is not None else 0,
                    row_data[21] if len(row_data) > 21 and row_data[21] is not None else 0,
                    row_data[22] if len(row_data) > 22 and row_data[22] is not None else 0,
                    row_data[23] if len(row_data) > 23 and row_data[23] is not None else 0,
                ],
            })

        print(f"Se extrajeron {len(flowers_data)} flores.")
    finally:
        try:
            graph_request("POST", f"{plan_workbook_url}/closeSession", psh, timeout=30)
        except Exception:
            pass

    if os.environ.get("SHAREPOINT_UPLOAD", "true").lower() not in {"1", "true", "yes", "si", "sí"}:
        print("Modo prueba: SHAREPOINT_UPLOAD esta apagado.")
        return

    print("\nIniciando escritura en Requerimiento vs proyeccion...")
    drive_id     = item_req["parentReference"]["driveId"]
    workbook_url = f"{GRAPH_URL}/drives/{drive_id}/items/{item_req['id']}/workbook"

    # -----------------------------------------------------------------------
    # SESION UNICA: leer, borrar con null, y escribir en orden fila por fila
    # -----------------------------------------------------------------------
    print("Abriendo sesion en DataProy...")
    sess = graph_request(
        "POST", f"{workbook_url}/createSession", headers,
        json={"persistChanges": True}, timeout=60,
    ).json()
    sh = {**headers, "workbook-session-id": sess["id"]}

    try:
        used      = graph_request("GET", f"{workbook_url}/worksheets/DataProy/usedRange", sh, timeout=120).json()
        address   = used.get("address", "")
        values    = used.get("values", [])

        m             = re.search(r'!([A-Za-z]+)(\d+)', address)
        start_col_str = m.group(1).upper() if m else "A"
        start_row     = int(m.group(2))    if m else 1
        start_col_idx = col_letter_to_index(start_col_str)
        col_h_rel     = col_letter_to_index("H") - start_col_idx

        # Encontrar donde empieza la zona CORTE (despues de COMPRA)
        corte_start_excel_row = None
        seen_compra = False
        for idx, row_data in enumerate(values):
            desc = str(row_data[col_h_rel]).strip().upper() if (col_h_rel >= 0 and len(row_data) > col_h_rel) else ""
            if desc == "COMPRA":
                seen_compra = True
            elif seen_compra and desc != "COMPRA":
                corte_start_excel_row = start_row + idx
                break

        if not corte_start_excel_row:
            corte_start_excel_row = start_row + len(values)

        num_flowers  = len(flowers_data)
        total_needed = num_flowers * len(weeks)

        rows_by_week = {w: [] for w in weeks}
        for i in range(total_needed):
            week_idx = i // num_flowers
            if week_idx < len(weeks):
                rows_by_week[weeks[week_idx]].append(corte_start_excel_row + i)

        # Filas residuales de ejecuciones anteriores
        end_row = corte_start_excel_row + total_needed
        rows_to_clear = []
        for idx, row_data in enumerate(values):
            excel_row = start_row + idx
            if excel_row >= end_row:
                desc = str(row_data[col_h_rel]).strip().upper() if (col_h_rel >= 0 and len(row_data) > col_h_rel) else ""
                if desc == "CORTE":
                    rows_to_clear.append(excel_row)

        # Escribir semana por semana: null-clear O&S → write F:H, O, S
        for i, week in enumerate(weeks):
            target_rows = sorted(rows_by_week[week])
            if not target_rows:
                print(f"Advertencia: No hay filas para semana {week}.")
                continue

            flower_idx = 0
            for block in make_blocks(target_rows):
                sr, er, count = block[0], block[-1], len(block)
                fcd_vals, tallos_vals, sem_vals = [], [], []

                for _ in range(count):
                    if flower_idx < len(flowers_data):
                        item = flowers_data[flower_idx]
                        fcd_vals.append([
                            str(item["flor"])  if item["flor"]  else "",
                            str(item["color"]) if item["color"] else "",
                            "CORTE",
                        ])
                        qty = item["qtys"][i]
                        try:
                            tallos_vals.append([int(qty) if qty is not None else 0])
                        except (ValueError, TypeError):
                            tallos_vals.append([0])
                    else:
                        fcd_vals.append(["", "", "CORTE"])
                        tallos_vals.append([0])
                    try:
                        sem_vals.append([int(week)])
                    except (ValueError, TypeError):
                        sem_vals.append([week])
                    flower_idx += 1

                print(f"Semana {week}: null-clear + write filas {sr}:{er}...")
                clear_and_write(workbook_url, sh, sr, er, count,
                                fcd_vals, tallos_vals, sem_vals)

        # Limpiar filas residuales viejas
        if rows_to_clear:
            print(f"Limpiando {len(rows_to_clear)} filas residuales...")
            for block in make_blocks(rows_to_clear):
                sr, er, cnt = block[0], block[-1], len(block)
                patch(workbook_url, sh, f"F{sr}:H{er}", [["", "", ""] for _ in range(cnt)])
                patch(workbook_url, sh, f"O{sr}:O{er}", [[None] for _ in range(cnt)])
                patch(workbook_url, sh, f"S{sr}:S{er}", [[None] for _ in range(cnt)])

        print("Escritura finalizada con exito.")
    finally:
        print("Cerrando sesion de Excel Online...")
        try:
            graph_request("POST", f"{workbook_url}/closeSession", sh, timeout=30)
        except Exception as exc:
            print(f"Aviso al cerrar sesion: {exc}")


if __name__ == "__main__":
    main()

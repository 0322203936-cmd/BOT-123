import os
import re
import sys
from pathlib import Path
import openpyxl

from sharepoint_sync import (
    GRAPH_URL,
    graph_headers,
    graph_token,
    resolve_sharepoint_item_by_url,
    download_sharepoint_file,
)
from excel_range_sync import graph_request

REQ_PROY_URL = "https://pacificafarms.sharepoint.com/:x:/r/sites/requerimientovsproyeccion/_layouts/15/Doc.aspx?sourcedoc=%7B277A76AA-508A-47F8-8A4A-F19D46660D65%7D&file=Requerimiento%20vs%20proyeccion%20Test.xlsm&action=default&mobileredirect=true"
PLAN_COSECHA_URL = "https://pacificafarms.sharepoint.com/:x:/r/sites/requerimientovsproyeccion/_layouts/15/Doc.aspx?sourcedoc=%7B0A3464AB-7BD8-400A-A0E6-5BC92E23CE3E%7D&file=Plan%20de%20cosecha%202026%20Test.xlsx&action=default&mobileredirect=true"

def col_letter_to_index(letter):
    idx = 0
    for char in letter.upper():
        idx = idx * 26 + (ord(char) - ord('A') + 1)
    return idx - 1

def main():
    print("Obteniendo token de Microsoft Graph...")
    token = graph_token()
    
    print("Resolviendo archivos en SharePoint...")
    item_req = resolve_sharepoint_item_by_url(token, REQ_PROY_URL)
    item_plan = resolve_sharepoint_item_by_url(token, PLAN_COSECHA_URL)
    
    print("Descargando Plan de cosecha (solo lectura)...")
    plan_file = download_sharepoint_file(token, item_plan, "Plan de cosecha 2026 Test_BOT.xlsx")
    
    print(f"Abriendo {plan_file}...")
    wb_plan = openpyxl.load_workbook(plan_file, data_only=True)
    
    cosecha_sheets = [s for s in wb_plan.sheetnames if s.startswith("P Cosecha ")]
    if not cosecha_sheets:
        print("Error: No se encontraron hojas de 'P Cosecha'")
        sys.exit(1)
        
    latest_sheet_name = sorted(cosecha_sheets)[-1]
    ws_plan = wb_plan[latest_sheet_name]
    print(f"Hoja detectada: {latest_sheet_name}")
    
    current_week = int(latest_sheet_name[-2:])
    weeks = [current_week + i for i in range(8)]
    print(f"Semanas a procesar: {weeks}")

    flowers_data = []
    for row in range(5, ws_plan.max_row + 1):
        flor_val = ws_plan.cell(row=row, column=1).value
        if flor_val is not None and str(flor_val).strip().lower() == "total":
            break
        if not flor_val:
            continue
            
        flor_real = ws_plan.cell(row=row, column=4).value
        color_real = ws_plan.cell(row=row, column=5).value
        
        if not flor_real:
            continue
            
        qty_30 = ws_plan.cell(row=row, column=6).value or 0
        qty_31 = ws_plan.cell(row=row, column=19).value or 0
        qty_32 = ws_plan.cell(row=row, column=20).value or 0
        qty_33 = ws_plan.cell(row=row, column=21).value or 0
        qty_34 = ws_plan.cell(row=row, column=22).value or 0
        qty_35 = ws_plan.cell(row=row, column=23).value or 0
        qty_36 = ws_plan.cell(row=row, column=24).value or 0
        qty_37 = ws_plan.cell(row=row, column=25).value or 0
        
        flowers_data.append({
            "flor": flor_real,
            "color": color_real,
            "qtys": [qty_30, qty_31, qty_32, qty_33, qty_34, qty_35, qty_36, qty_37]
        })

    print(f"Se extrajeron {len(flowers_data)} flores.")
    wb_plan.close()

    if os.environ.get("SHAREPOINT_UPLOAD", "true").lower() not in {"1", "true", "yes", "si", "sí"}:
        print("Modo prueba: SHAREPOINT_UPLOAD está apagado. Deteniendo ejecución.")
        return

    print("\nIniciando Edición en Vivo (Live Patching) en Requerimiento vs proyeccion...")
    drive_id = item_req["parentReference"]["driveId"]
    workbook_url = f"{GRAPH_URL}/drives/{drive_id}/items/{item_req['id']}/workbook"
    headers = {**graph_headers(token), "Content-Type": "application/json"}
    
    print("Abriendo sesión en Excel Online...")
    session = graph_request(
        "POST",
        f"{workbook_url}/createSession",
        headers,
        json={"persistChanges": True},
        timeout=60,
    ).json()
    session_headers = {**headers, "workbook-session-id": session["id"]}
    
    try:
        print("Obteniendo área de trabajo (usedRange)...")
        used_range_res = graph_request(
            "GET", 
            f"{workbook_url}/worksheets/DataProy/usedRange", 
            session_headers,
            timeout=120
        ).json()
        
        address = used_range_res.get("address", "")
        values = used_range_res.get("values", [])
        
        match = re.search(r'!([A-Za-z]+)(\d+)', address)
        start_col_str = match.group(1).upper() if match else "A"
        start_row = int(match.group(2)) if match else 1
        
        start_col_idx = col_letter_to_index(start_col_str)
        col_h_rel = col_letter_to_index("H") - start_col_idx
        col_s_rel = col_letter_to_index("S") - start_col_idx
        
        rows_by_week = {w: [] for w in weeks}
        
        in_corte_block = False
        corte_count = 0
        num_flowers = len(flowers_data)
        
        for idx, row_data in enumerate(values):
            if len(row_data) > col_h_rel and col_h_rel >= 0:
                desc = str(row_data[col_h_rel]).strip().upper()
                
                if desc == "CORTE" or (desc in ["", "NONE", "NULL"] and in_corte_block):
                    in_corte_block = True
                    
                    if num_flowers > 0:
                        week_idx = corte_count // num_flowers
                        if week_idx < len(weeks):
                            current_week = weeks[week_idx]
                            excel_row = start_row + idx
                            rows_by_week[current_week].append(excel_row)
                            
                    corte_count += 1
                else:
                    in_corte_block = False
                    
        for i, week in enumerate(weeks):
            target_rows = sorted(rows_by_week[week])
            if not target_rows:
                print(f"Advertencia: No se encontraron filas CORTE para semana {week}.")
                continue
                
            blocks = []
            current_block = []
            for r in target_rows:
                if not current_block or r == current_block[-1] + 1:
                    current_block.append(r)
                else:
                    blocks.append(current_block)
                    current_block = [r]
            if current_block:
                blocks.append(current_block)
                
            flower_idx = 0
            for block in blocks:
                start_r = block[0]
                end_r = block[-1]
                count = len(block)
                
                flor_color_desc_values = []
                tallos_values = []
                semana_values = []
                
                for _ in range(count):
                    if flower_idx < len(flowers_data):
                        item = flowers_data[flower_idx]
                        flor_color_desc_values.append([item["flor"], item["color"], "CORTE"])
                        tallos_values.append([item["qtys"][i]])
                    else:
                        flor_color_desc_values.append(["", "", "CORTE"])
                        tallos_values.append([""])
                    semana_values.append([week])
                    flower_idx += 1
                    
                print(f"Semana {week}: Escribiendo bloque filas {start_r}:{end_r}...")
                address_fh = f"F{start_r}:H{end_r}"
                graph_request(
                    "PATCH", 
                    f"{workbook_url}/worksheets/DataProy/range(address='{address_fh}')", 
                    session_headers, 
                    json={"values": flor_color_desc_values}
                )
                
                address_o = f"O{start_r}:O{end_r}"
                graph_request(
                    "PATCH", 
                    f"{workbook_url}/worksheets/DataProy/range(address='{address_o}')", 
                    session_headers, 
                    json={"values": tallos_values}
                )
                
                address_s = f"S{start_r}:S{end_r}"
                graph_request(
                    "PATCH", 
                    f"{workbook_url}/worksheets/DataProy/range(address='{address_s}')", 
                    session_headers, 
                    json={"values": semana_values}
                )

        print("Edición en vivo finalizada con éxito.")
    finally:
        print("Cerrando sesión de Excel Online...")
        try:
            graph_request(
                "POST",
                f"{workbook_url}/closeSession",
                session_headers,
                timeout=30,
            )
        except Exception as exc:
            print(f"Aviso al cerrar la sesion: {exc}")

if __name__ == "__main__":
    main()

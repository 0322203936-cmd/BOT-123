import os
import sys
from pathlib import Path
import openpyxl

from sharepoint_sync import (
    graph_token,
    resolve_sharepoint_item_by_url,
    download_sharepoint_file,
    upload_sharepoint_file
)

REQ_PROY_URL = "https://pacificafarms.sharepoint.com/:x:/r/sites/requerimientovsproyeccion/_layouts/15/Doc.aspx?sourcedoc=%7B277A76AA-508A-47F8-8A4A-F19D46660D65%7D&file=Requerimiento%20vs%20proyeccion%20Test.xlsm&action=default&mobileredirect=true"
PLAN_COSECHA_URL = "https://pacificafarms.sharepoint.com/:x:/r/sites/requerimientovsproyeccion/_layouts/15/Doc.aspx?sourcedoc=%7B0A3464AB-7BD8-400A-A0E6-5BC92E23CE3E%7D&file=Plan%20de%20cosecha%202026%20Test.xlsx&action=default&mobileredirect=true"

def main():
    print("Obteniendo token de Microsoft Graph...")
    token = graph_token()
    
    print("Resolviendo archivos en SharePoint...")
    item_req = resolve_sharepoint_item_by_url(token, REQ_PROY_URL)
    item_plan = resolve_sharepoint_item_by_url(token, PLAN_COSECHA_URL)
    
    print("Descargando Plan de cosecha...")
    plan_file = download_sharepoint_file(token, item_plan, "Plan de cosecha 2026 Test_BOT.xlsx")
    
    print("Descargando Requerimiento vs proyeccion...")
    req_file = download_sharepoint_file(token, item_req, "Requerimiento vs proyeccion Test_BOT.xlsm")

    print(f"Abriendo {plan_file} (solo lectura de datos)...")
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

    print(f"Abriendo {req_file} para escritura (conservando VBA y fórmulas)...")
    wb_req = openpyxl.load_workbook(req_file, keep_vba=True)
    if "DataProy" not in wb_req.sheetnames:
        print("Error: No se encontró la hoja DataProy")
        sys.exit(1)
        
    ws_req = wb_req["DataProy"]
    
    rows_by_week = {w: [] for w in weeks}
    for r in range(2, ws_req.max_row + 1):
        desc = ws_req.cell(row=r, column=4).value
        semana = ws_req.cell(row=r, column=15).value
        if desc == "CORTE" and semana in weeks:
            rows_by_week[semana].append(r)
            
    for i, week in enumerate(weeks):
        target_rows = rows_by_week[week]
        if not target_rows:
            print(f"Advertencia: No se encontraron filas pre-creadas para CORTE en la semana {week}.")
            continue
            
        print(f"Semana {week}: pegando {len(flowers_data)} registros en {len(target_rows)} filas disponibles.")
        
        for idx, row_num in enumerate(target_rows):
            if idx < len(flowers_data):
                item = flowers_data[idx]
                ws_req.cell(row=row_num, column=2).value = item["flor"]
                ws_req.cell(row=row_num, column=3).value = item["color"]
                ws_req.cell(row=row_num, column=11).value = item["qtys"][i]
            else:
                ws_req.cell(row=row_num, column=2).value = None
                ws_req.cell(row=row_num, column=3).value = None
                ws_req.cell(row=row_num, column=11).value = None

    print("Guardando cambios localmente...")
    wb_req.save(req_file)
    wb_req.close()
    
    if os.environ.get("SHAREPOINT_UPLOAD", "true").lower() in {"1", "true", "yes", "si", "sí"}:
        print("Subiendo Requerimiento vs proyeccion a SharePoint...")
        upload_sharepoint_file(token, item_req, req_file)
    else:
        print("Modo de prueba: El archivo fue modificado pero SHAREPOINT_UPLOAD está apagado. No se subió a SharePoint.")

    print("Proceso completado exitosamente.")

if __name__ == "__main__":
    main()

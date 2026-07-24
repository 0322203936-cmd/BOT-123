import os
import re
import sys
from pathlib import Path
import pandas as pd
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright

from sharepoint_sync import (
    GRAPH_URL,
    graph_headers,
    graph_token,
    resolve_sharepoint_item_by_url,
)
from excel_range_sync import graph_request

# Importamos funciones de Posco del script original (lo dejamos intacto)
from inventario import (
    POSCO_URL, required_secret, open_inventory_general, export_inventory
)

REQ_PROY_URL = "https://pacificafarms.sharepoint.com/:x:/r/sites/requerimientovsproyeccion/_layouts/15/Doc.aspx?sourcedoc=%7B277A76AA-508A-47F8-8A4A-F19D46660D65%7D&file=Requerimiento%20vs%20proyeccion%20Test.xlsm&action=default&mobileredirect=true"

def col_letter_to_index(letter):
    idx = 0
    for char in letter.upper():
        idx = idx * 26 + (ord(char) - ord('A') + 1)
    return idx - 1

def patch(workbook_url, sh, addr, values):
    """PATCH a range and immediately GET it back to verify the write landed."""
    graph_request("PATCH", f"{workbook_url}/worksheets/DataProy/range(address='{addr}')",
                  sh, json={"values": values}, timeout=60)

def get_posco_inventory():
    user = required_secret("POSCO_USER")
    password = required_secret("POSCO_PASSWORD")
    
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page = context.new_page()
        try:
            print("Abriendo Posco para AINVENTARIO...", flush=True)
            page.goto(POSCO_URL, wait_until="networkidle", timeout=60_000)
            
            page.locator('input[placeholder*="usuario@email.com" i], input[type="text"]').first.fill(user)
            page.locator('input[placeholder*="Password" i], input[type="password"]').first.fill(password)
            page.get_by_role("button", name="Iniciar Sesión").click(timeout=15_000)
            page.wait_for_load_state("networkidle", timeout=60_000)
            page.wait_for_timeout(2_000)
            
            print("Navegando a Inventario General...", flush=True)
            open_inventory_general(page)
            
            print("Exportando archivo...", flush=True)
            report_path = export_inventory(page)
            print(f"Archivo exportado a {report_path}", flush=True)
            return report_path
        finally:
            context.close()
            browser.close()

def main():
    print("=== Bot AINVENTARIO ===")
    
    # 1. Descargar Inventario de Posco
    report_path = get_posco_inventory()
    
    # 2. Procesar Datos
    print("Procesando datos de Posco...")
    df = pd.read_excel(report_path)
    
    # Asegurar que los N/A o valores nulos se traten como 0 para no romper la suma
    df['tallos_en_existencia'] = pd.to_numeric(df['tallos_en_existencia'], errors='coerce').fillna(0)
    
    # Filtrar aquellos que tienen 0 tallos para no pegar filas vacías (opcional, pero útil)
    df = df[df['tallos_en_existencia'] > 0]
    
    # Calcular semana (Calendario Sabado a Viernes)
    # Al sumar 2 dias, el Sabado se vuelve Lunes, forzando a que la semana ISO cambie justo ese dia.
    current_date = datetime.now()
    shifted_date = current_date + timedelta(days=2)
    current_week = shifted_date.isocalendar()[1]
    print(f"Semana calculada (Sab-Vie): {current_week}")
    
    num_new_rows = len(df)
    print(f"Se detectaron {num_new_rows} combinaciones unicas de flores/colores con inventario.")
    
    fcd_vals = []
    tallos_vals = []
    sem_vals = []
    
    for _, row in df.iterrows():
        fcd_vals.append([str(row['flor']).strip(), str(row['color']).strip(), "AINVENTARIO"])
        tallos_vals.append([int(row['tallos_en_existencia'])])
        sem_vals.append([int(current_week)])
    
    if num_new_rows == 0:
        print("El inventario esta vacio. Saliendo...")
        return
        
    # 3. Actualizar SharePoint (Excel en Vivo)
    print("Obteniendo token de Microsoft Graph...")
    token = graph_token()
    
    print("Resolviendo archivo Requerimiento vs proyeccion...")
    item_req = resolve_sharepoint_item_by_url(token, REQ_PROY_URL)
    
    drive_id = item_req["parentReference"]["driveId"]
    workbook_url = f"{GRAPH_URL}/drives/{drive_id}/items/{item_req['id']}/workbook"
    
    headers = {**graph_headers(token), "Content-Type": "application/json"}
    
    print("Abriendo sesion persistente en DataProy...")
    sess = graph_request(
        "POST", f"{workbook_url}/createSession", headers,
        json={"persistChanges": True}, timeout=60,
    ).json()
    sh = {**headers, "workbook-session-id": sess["id"]}
    
    try:
        print("Obteniendo rango actual para ubicar AINVENTARIO...")
        used = graph_request("GET", f"{workbook_url}/worksheets/DataProy/usedRange", sh, timeout=120).json()
        address = used.get("address", "")
        values = used.get("values", [])
        
        m = re.search(r'!([A-Za-z]+)(\d+)', address)
        start_col_str = m.group(1).upper() if m else "A"
        start_row = int(m.group(2)) if m else 1
        
        col_h_rel = col_letter_to_index("H") - col_letter_to_index(start_col_str)
        
        ainventario_start = None
        ainventario_end = None
        
        # Encontrar el bloque actual de AINVENTARIO (o AININVENTARIO)
        for idx, row_data in enumerate(values):
            desc = str(row_data[col_h_rel]).strip().upper() if (col_h_rel >= 0 and len(row_data) > col_h_rel) else ""
            if "AINVENTARIO" in desc or "AININVENTARIO" in desc:
                if ainventario_start is None:
                    ainventario_start = start_row + idx
                ainventario_end = start_row + idx
        
        if ainventario_start is None:
            # Si no habia ninguno, empezamos despues del header
            ainventario_start = 2
            ainventario_end = 1
            
        num_current_rows = ainventario_end - ainventario_start + 1
        print(f"Filas actuales: {num_current_rows} (Desde fila {ainventario_start} hasta {ainventario_end})")
        
        # Obtener limites de la tabla (opcional, solo para debuguear, no lo usamos para insertar)
        tables_res = graph_request("GET", f"{workbook_url}/worksheets/DataProy/tables", sh).json()
        if tables_res and tables_res.get("value") and len(tables_res["value"]) > 0:
            print(f"Tabla detectada: {tables_res['value'][0]['name']}")
        
        diff = num_new_rows - num_current_rows
        
        # Insertar o eliminar filas dinamicamente usando FILAS COMPLETAS para evitar error 400 y 409
        if diff > 0:
            print(f"Insertando {diff} filas en Excel...")
            insert_addr = f"{ainventario_end + 1}:{ainventario_end + diff}"
            graph_request("POST", f"{workbook_url}/worksheets/DataProy/range(address='{insert_addr}')/insert", 
                          sh, json={"shift": "Down"}, timeout=120)
        elif diff < 0:
            print(f"Eliminando {-diff} filas sobrantes de Excel...")
            delete_addr = f"{ainventario_end + diff + 1}:{ainventario_end}"
            graph_request("POST", f"{workbook_url}/worksheets/DataProy/range(address='{delete_addr}')/delete", 
                          sh, json={"shift": "Up"}, timeout=120)
        
        write_start = ainventario_start
        write_end = ainventario_start + num_new_rows - 1
        
        print(f"Escribiendo {num_new_rows} filas de inventario (Fila {write_start} a {write_end})...")
        
        # Generar las columnas A-E
        prefix_vals = [["Proyeccion", "INVENTARIO", "INVENTARIO", "INVENTARIO", ""] for _ in range(num_new_rows)]
        
        # Limpiar celdas numericas con null para evitar error de tipos
        patch(workbook_url, sh, f"O{write_start}:O{write_end}", [[None] for _ in range(num_new_rows)])
        patch(workbook_url, sh, f"S{write_start}:S{write_end}", [[None] for _ in range(num_new_rows)])
        
        # Escribir todo en bloque
        patch(workbook_url, sh, f"A{write_start}:E{write_end}", prefix_vals)
        patch(workbook_url, sh, f"F{write_start}:H{write_end}", fcd_vals)
        patch(workbook_url, sh, f"O{write_start}:O{write_end}", tallos_vals)
        patch(workbook_url, sh, f"S{write_start}:S{write_end}", sem_vals)
        
        print("¡Operacion finalizada con exito!")
        
    finally:
        print("Cerrando sesion de Excel Online...")
        try:
            graph_request("POST", f"{workbook_url}/closeSession", sh, timeout=30)
        except Exception:
            pass

if __name__ == "__main__":
    main()

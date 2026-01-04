import os
import re
import pandas as pd
import requests
from dotenv import load_dotenv #
from playwright.sync_api import sync_playwright

# ==========================================
# 1. CONFIGURACIÓN DE RUTAS Y SEGURIDAD
# ==========================================
# Detecta la raíz del proyecto desde 'src'
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Carga las variables del archivo .env que está en la raíz
load_dotenv(os.path.join(BASE_DIR, ".env")) 
USER_ID_BB = os.getenv("USER_ID_BB") #

# Carpeta de destino 00_inputs
CARPETA_DATA = os.path.join(BASE_DIR, "01_data")
os.makedirs(CARPETA_DATA, exist_ok=True)
ARCHIVO_SALIDA = os.path.join(CARPETA_DATA, "base_maestra_ids.xlsx")

def run():
    if not USER_ID_BB:
        print("❌ ERROR: No se encontró USER_ID_BB en el archivo .env")
        return

    with sync_playwright() as p:
        print(f"--- 🎭 INICIANDO PLAYWRIGHT ---")
        
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        page.goto("https://upn-colaborador.blackboard.com/")
        print("\n🔑 ESPERANDO LOGIN MANUAL...")
        input("✅ Cuando veas tus cursos, presiona ENTER aquí...")

        cookies = context.cookies()
        cookie_string = "; ".join([f"{c['name']}={c['value']}" for c in cookies])
        browser.close()

        # ==========================================
        # 2. CONSUMO DE API
        # ==========================================
        url = f"https://upn.blackboard.com/learn/api/v1/users/{USER_ID_BB}/memberships?expand=course.effectiveAvailability,course.permissions,courseRole&includeCount=true&limit=10000"
        
        headers = {
            "Cookie": cookie_string,
            "User-Agent": "Mozilla/5.0",
            "Content-Type": "application/json"
        }

        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            data = response.json()
            lista_cursos = []

            for item in data.get('results', []):
                curso_obj = item.get('course', {})
                nombre_full = curso_obj.get('name', '')
                match = re.search(r'(\d{6}\.\d{4})', nombre_full)
                id_limpio = match.group(1) if match else "N/A"

                lista_cursos.append({
                    "ID": id_limpio,
                    "Nombre": nombre_full,
                    "ID_Interno": curso_obj.get('id'),
                    "ID_Visible": curso_obj.get('courseId')
                })

            df = pd.DataFrame(lista_cursos)
            df = df[["ID", "Nombre", "ID_Interno", "ID_Visible"]]

            # ==========================================
            # 3. EXPORTACIÓN CON XLSXWRITER
            # ==========================================
            writer = pd.ExcelWriter(ARCHIVO_SALIDA, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Mapa')
            
            workbook  = writer.book
            worksheet = writer.sheets['Mapa']
            formato_texto = workbook.add_format({'num_format': '@'})
            
            worksheet.set_column('A:A', 20, formato_texto) 
            worksheet.set_column('B:B', 70)                
            worksheet.set_column('C:C', 25)                
            worksheet.set_column('D:D', 40)                
            
            writer.close()
            print(f"✨ ¡Éxito! Archivo generado en: {ARCHIVO_SALIDA}")
        else:
            print(f"❌ Error API: {response.status_code}")

if __name__ == "__main__":
    run()
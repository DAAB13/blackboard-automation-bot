import time
import pandas as pd
import requests
import os
from datetime import datetime
from seleniumwire import webdriver # captura respuestas de manera sileciosa
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ==========================================
# 1. CONFIGURACIÓN DE RUTAS (ADAPTADO A TU ESTRUCTURA)
# ==========================================
# Calculamos la ruta base del proyecto (subiendo un nivel desde 'scr')
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Input: Viene de la carpeta 01_data (donde lo dejó el script anterior)
ARCHIVO_INPUT = os.path.join(BASE_DIR, "01_data", "resumen_con_llave.xlsx")

# Output: Va a la carpeta 02_outputs (Reporte Final)
ARCHIVO_SALIDA = os.path.join(BASE_DIR, "02_outputs", "REPORTE_FINAL_COMPLETO.xlsx")

print("--- 🤖 ROBOT UPN: MODO PRODUCCIÓN (TODOS LOS CURSOS) ---")

if not os.path.exists(ARCHIVO_INPUT):
    print(f"❌ Error CRÍTICO: No encuentro el archivo de entrada en:\n{ARCHIVO_INPUT}")
    print("👉 Ejecuta primero '01_etl_programacion.py'")
    exit()

# Leemos forzando que los IDs sean texto (str) para no perder ceros a la izquierda
df_trabajo = pd.read_excel(ARCHIVO_INPUT, dtype={'ID': str, 'ID_Interno': str})

# --- SIN FRENOS ---
print(f"📚 Cursos detectados en archivo: {len(df_trabajo)}")
print("🚀 Iniciando procesamiento masivo...")

# --- DETECTIVE DE COLUMNAS (Para encontrar el nombre del curso automáticamente) ---
col_curso = 'ND'
posibles_nombres = ['Curso', 'Nombre', 'Asignatura', 'Materia', 'Descripción']
for col in df_trabajo.columns:
    for posible in posibles_nombres:
        if posible.lower() in col.lower():
            col_curso = col
            break
    if col_curso != 'ND': break
print(f"👉 Nombre del curso tomado de columna: '{col_curso}'")

# ==========================================
# 2. INICIAR NAVEGADOR (CON PROXY INTERNO)
# ==========================================
options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors') # le dice a chrome que no se detenga si se encuentra una alerta de red (comun en redes corporativas)
options.set_capability("acceptInsecureCerts", True)
options.add_argument("--start-maximized")

# Iniciamos Chrome con selenium-wire para capturar tokens
# webdriver.Chrome abre el navegador, service conecta python con el navegador físico
# ChromeDriverManager().install() chrome se actualiza automaticamente
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options) 

# 3. LOGIN MANUAL
driver.get("https://upn-colaborador.blackboard.com/")
print("\n🔑 POR FAVOR, INICIA SESIÓN MANUALMENTE...")
input("👉 Presiona ENTER en esta consola cuando ya veas la lista de tus cursos...")

# ==========================================
# 4. EXTRACCIÓN MASIVA
# ==========================================
lista_final = []

for index, fila in df_trabajo.iterrows(): # iterrows recorre fila por fila
    id_nrc = fila.get('ID')          
    id_nav = fila.get('ID_Interno')
    nombre_curso_excel = fila.get(col_curso, 'ND') # busca la columna col_curso, si no pornle nd y sigue trabajando
    
    # Validación básica: si no hay ID interno, saltamos
    if pd.isna(id_nav) or id_nav == "" or id_nav == "nan": 
        continue

    print(f"[{index+1}/{len(df_trabajo)}] Procesando NRC: {id_nrc}.  ", end=" ")

    try:
        # Limpiamos peticiones anteriores
        del driver.requests
        
        # Navegamos directo a la sección de grabaciones
        driver.get(f"https://upn.blackboard.com/ultra/courses/{id_nav}/outline/collab/launchRecordings")
        
        token_encontrado = None
        # Esperamos máx 15 seg por el token
        for _ in range(15):
            time.sleep(1) # Espera leve
            for request in driver.requests: # revisa uno por uno
                if request.response and "bbcollab.com" in request.url and "recordings" in request.url:
                    auth = request.headers.get('Authorization') # Authorization es la llave maestra
                    if auth and "Bearer" in auth: # Bearer token significa el q tiene la llave tiene permiso
                        token_encontrado = auth
                        break
            if token_encontrado: break
        
        if token_encontrado:
            # Consulta a la API de Collaborate (trae hasta 500 videos desde 2024)
            api_url = "https://us-lti.bbcollab.com/collab/api/csa/recordings?startTime=2024-01-01T00:00:00.000Z&limit=500"
            resp = requests.get(api_url, headers={"Authorization": token_encontrado}, timeout=10)
            
            if resp.status_code == 200:
                data = resp.json().get('results', [])
                print(f" -> ✅ {len(data)} videos.")
                
                for v in data:
                    fecha_cruda = v.get('startTime')
                    try:
                        # Convertimos a objeto datetime para poder manipularlo luego
                        dt_obj = pd.to_datetime(fecha_cruda).replace(tzinfo=None)
                        solo_hora = dt_obj.time().replace(microsecond=0)
                    except:
                        dt_obj = None
                        solo_hora = "00:00:00"

                    lista_final.append({
                        'ID': id_nrc,
                        'Curso': nombre_curso_excel,
                        'Docente': fila.get('DOCENTE', fila.get('Profesor', '')),
                        'Nombre Video': v.get('mediaName'),
                        'Fecha': dt_obj, # IMPORTANTE: Se guarda como OBJETO FECHA
                        'Hora': solo_hora,
                        'Duración (min)': round(v.get('duration', 0) / 60000, 1),
                        'Link': v.get('guestLink') or f"https://us.bbcollab.com/recording/{v.get('id')}"
                    })
            else: 
                print(f" -> ⚠️ Error API: {resp.status_code}")
        else: 
            print(" -> ❌ Sin Token (Curso sin grabaciones o error de carga).")

    except Exception as e:
        print(f" -> ❌ Error general: {e}")

# Cerramos navegador al terminar todo
driver.quit()

# ==========================================
# 5. EXPORTACIÓN TIPO "SUPERVISIÓN" (MERGE READY)
# ==========================================
if lista_final:
    print("\n>>> Generando reporte final optimizado...")
    df_export = pd.DataFrame(lista_final)

    # LIMPIEZA FINAL DE FECHA:
    # .normalize() elimina la hora interna (la pone en 00:00:00) para que cruce perfecto con tu panel.
    df_export['Fecha'] = pd.to_datetime(df_export['Fecha'], errors='coerce').dt.normalize()

    # Seleccionamos columnas en orden
    cols = ['ID', 'Curso', 'Docente', 'Nombre Video', 'Fecha', 'Hora', 'Duración (min)', 'Link']
    df_export = df_export[[c for c in cols if c in df_export.columns]]

    # EXPORTACIÓN CON FORMATO AVANZADO (XLSXWRITER)
    # datetime_format='dd/mm/yyyy' fuerza la vista visual correcta
    with pd.ExcelWriter(ARCHIVO_SALIDA, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
        df_export.to_excel(writer, sheet_name='Reporte', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Reporte']
        
        formato_centrar = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        
        # A) Convertir rango en TABLA INTELIGENTE (Estilo Azul)
        (max_fila, max_col) = df_export.shape
        lista_columnas = [{'header': col} for col in df_export.columns]
        
        worksheet.add_table(0, 0, max_fila, max_col - 1, {
            'columns': lista_columnas,
            'style': 'TableStyleMedium2',
            'name': 'TablaReporteCompleto'
        })
        
        # B) Ajustar anchos de columnas
        worksheet.set_column(0, max_col - 1, 15, formato_centrar) # Ancho base centrado
        worksheet.set_column('B:B', 30) # Curso
        worksheet.set_column('C:C', 35) # Docente (Aumentado un poco por nombres largos)
        worksheet.set_column('E:E', 15) # Fecha (formato controlado por writer)
        worksheet.set_column('H:H', 65) # Link
        
    print(f"------------------------------------------------")
    print(f"✅ PROCESO COMPLETADO.")
    print(f"   Archivo generado: {ARCHIVO_SALIDA}")
    print(f"   Total de registros: {len(df_export)}")
    print(f"   Formato Fecha: dd/mm/yyyy (Visual) | Value (Date) -> Listo para Merge")
    print(f"------------------------------------------------")

else:
    print("\n⚠️ ALERTA: No se extrajeron datos de ningún curso.")
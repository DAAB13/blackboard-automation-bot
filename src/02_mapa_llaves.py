import time
import requests
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service #el motor de chorme arranque correctamente en el sistema operativo
from webdriver_manager.chrome import ChromeDriverManager

# ==========================================
# 1. CONFIGURACIÓN DE RUTAS
# ==========================================
# Subimos un nivel (..) para salir de 'scr' y entrar a '01_data'
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARCHIVO_SALIDA = os.path.join(BASE_DIR, "01_data", "base_maestra_ids.xlsx")

# ID de Usuario en Blackboard (Extraído de tu URL original: _567444_1)
# OJO: Si este ID cambia por usuario, avísame para automatizar su extracción también.
USER_ID_BB = "_567444_1" 

print("--- 🗺️ MAPA DE LLAVES: MODO HÍBRIDO (SELENIUM + API) ---")

# ==========================================
# 2. OBTENER COOKIE AUTOMÁTICAMENTE (Selenium)
# ==========================================
print("\n>>> Paso 1: Iniciando navegador para capturar sesión...")
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument('--ignore-certificate-errors')

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()), # actualiza la versión de chrome correcta para el script
    options=options
)

try:
    # A. Login Manual
    driver.get("https://upn-colaborador.blackboard.com/")
    print("🔑 POR FAVOR, INICIA SESIÓN MANUALMENTE...")
    print("👉 Esperamos a que cargue la página principal de Blackboard.")
    
    # B. Esperar a que el usuario presione ENTER (asegurando login completo)
    input("✅ Una vez dentro (viendo tus cursos), presiona ENTER aquí en la consola...")

    # C. Capturar Cookies del navegador
    print("🍪 Extrayendo cookies de la sesión...")
    selenium_cookies = driver.get_cookies()
    
    # Formatear cookies para la librería 'requests'
    # Las unimos en un solo string: "nombre=valor; nombre2=valor2"
    cookie_string = "; ".join([f"{c['name']}={c['value']}" for c in selenium_cookies])

    # Ya no necesitamos el navegador, cerramos para liberar RAM
    driver.quit()
    print("✅ Cookies capturadas con éxito. Cerrando navegador.")

    # ==========================================
    # 3. CONSUMO DE API (Tu lógica original)
    # ==========================================
    print("\n>>> Paso 2: Consultando API de Blackboard...")
    
    url = f"https://upn.blackboard.com/learn/api/v1/users/{USER_ID_BB}/memberships?expand=course.effectiveAvailability,course.permissions,courseRole&includeCount=true&limit=10000"

    headers = {
        "Cookie": cookie_string, # Usamos la cookie fresca de Selenium
        "User-Agent": "Mozilla/5.0", # para que el servidor piense que es un humano
        "Content-Type": "application/json" #
    }

    response = requests.get(url, headers=headers)
    
    if response.status_code == 200: # código universal
        data = response.json() # traducción
        lista_unica = [] # contenedor
        
        print(f"   Datos recibidos. Procesando {len(data.get('results', []))} registros...")

        for item in data.get('results', []):
            info_curso = item.get('course', {})
            id_vis = str(info_curso.get('courseId', ''))
            
            # --- TU LÓGICA DE LIMPIEZA (ID: 2025.02.225832.1049 -> 225832.1049) ---
            partes = id_vis.split('.')
            id_limpio = ""
            if len(partes) >= 4:
                # Tomas la parte 2 y 3 (índices 2 y 3)
                id_limpio = f"{partes[2]}.{partes[3]}"
            else:
                # Si el formato es raro, guardamos el original por seguridad
                id_limpio = id_vis
            
            lista_unica.append({
                'ID': str(id_limpio), 
                'Nombre': info_curso.get('displayName'),
                'ID_Interno': info_curso.get('id'), # La llave maestra (_123_1)
                'ID_Visible': id_vis
            })
        
        # Creamos el dataframe y eliminamos duplicados basados en ID_Interno
        df = pd.DataFrame(lista_unica).drop_duplicates(subset=['ID_Interno'])
        
        # ==========================================
        # 4. EXPORTACIÓN ROBUSTA (XlsxWriter)
        # ==========================================
        print(f"\n>>> Paso 3: Guardando Excel en {ARCHIVO_SALIDA}...")
        
        writer = pd.ExcelWriter(ARCHIVO_SALIDA, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Mapa')
        
        workbook  = writer.book # ingreo a las funciones avanzadas de excel
        worksheet = writer.sheets['Mapa']
        
        # DEFINIMOS EL FORMATO TEXTO (El @ es la clave en Excel)
        formato_texto = workbook.add_format({'num_format': '@'}) # num_format es una máscara de visualización
        
        # Aplicamos el formato y anchos
        worksheet.set_column('A:A', 20, formato_texto) # Columna ID forzada a Texto
        worksheet.set_column('B:B', 60)                
        worksheet.set_column('C:C', 25)                
        worksheet.set_column('D:D', 40)                
        
        writer.close()
        
        print(f"--------------------------------------------------")
        print(f"✅ ¡LISTO! Mapa generado exitosamente.")
        print(f"📂 Archivo: 01_data/base_maestra_ids.xlsx")
        print(f"--------------------------------------------------")
        
    else:
        print(f"❌ Error API {response.status_code}: La cookie no funcionó o el UserID es incorrecto.")

except Exception as e:
    print(f"❌ Error crítico: {e}")
    # Aseguramos cerrar driver si falló algo antes
    try: driver.quit() 
    except: pass
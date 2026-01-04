import pandas as pd
import os
import shutil
import time

# ==========================================
# 1. CONFIGURACIÓN DE RUTAS
# ==========================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

FILE_MAPA_IDS = os.path.join(BASE_DIR, "01_data", "base_maestra_ids.xlsx")
DIR_INPUTS = os.path.join(BASE_DIR, "00_inputs")
DIR_DATA = os.path.join(BASE_DIR, "01_data")
DIR_OUTPUTS = os.path.join(BASE_DIR, "02_outputs")

NOMBRE_ARCHIVO_PROG = "PANEL DE PROGRAMACIÓN V7.xlsx"
RUTA_ORIGEN_ONEDRIVE = fr"C:\Users\Diego AB\OneDrive - EduCorpPERU\POSGRADO-EPEC - Panel de Control Integrado\{NOMBRE_ARCHIVO_PROG}"
RUTA_TRABAJO_LOCAL = os.path.join(DIR_INPUTS, NOMBRE_ARCHIVO_PROG)

ARCHIVO_SUPERVISAR = os.path.join(DIR_DATA, "supervisar_clases.xlsx")
ARCHIVO_RESUMEN_LLAVE = os.path.join(DIR_DATA, "resumen_con_llave.xlsx")
ARCHIVO_ALERTAS = os.path.join(DIR_OUTPUTS, "reporte_alertas.xlsx")

print("--- 🧠 ETL: LIMPIEZA + GENERACIÓN DE REPORTES (CON FILTRO DE ACTIVOS) ---")

# ==========================================
# 2. COPIA DE SEGURIDAD
# ==========================================
print(f"\n>>> Paso 1: Copiando '{NOMBRE_ARCHIVO_PROG}'...")
if os.path.exists(RUTA_ORIGEN_ONEDRIVE):
    try:
        shutil.copy2(RUTA_ORIGEN_ONEDRIVE, RUTA_TRABAJO_LOCAL) # copy2 actua sobre los metadatos
        print("✅ Copia exitosa.")
    except PermissionError:
        print("⚠️ Archivo en uso. Intentando lectura directa del original...")
        RUTA_TRABAJO_LOCAL = RUTA_ORIGEN_ONEDRIVE # fallback
else:
    print(f"❌ NO SE ENCONTRÓ EL ARCHIVO EN ONEDRIVE:\n{RUTA_ORIGEN_ONEDRIVE}")
    exit()

# ==========================================
# 3. LÓGICA DE PROCESAMIENTO
# ==========================================
print("\n>>> Paso 2: Ejecutando lógica de limpieza...")

try:
    df_total = pd.read_excel(RUTA_TRABAJO_LOCAL, sheet_name='PROGRAMACIÓN', header=0, engine='openpyxl')

    columnas_interes = ['SOPORTE', 'CURSO', 'PERIODO', 'NRC', 'DOCENTE', 'SESIÓN', 'FECHAS', 'HORARIO', 'ESTADO DE CLASE']
    cols_existentes = [c for c in columnas_interes if c in df_total.columns] # filtro de seguridad para q no se rompa por si se cambia el nombre de una columna
    df_seguimiento = df_total[cols_existentes].copy()

    if 'PERIODO' in df_seguimiento.columns and 'NRC' in df_seguimiento.columns:
        df_seguimiento['ID'] = df_seguimiento['PERIODO'].astype(str) + '.' + df_seguimiento['NRC'].astype(str)

    if 'SOPORTE' in df_seguimiento.columns:
        df_diego = df_seguimiento[df_seguimiento['SOPORTE'].str.strip() == 'DIEGO'].copy()
    else:
        print("⚠️ Advertencia: Columna SOPORTE no encontrada.")
        exit()

    if not df_diego.empty:
        df_diego[['HORA_INI_STR', 'HORA_FIN_STR']] = df_diego['HORARIO'].str.split(' - ', expand=True) # expand=True crea dos columnas
        df_diego['HORA_INICIO'] = pd.to_datetime(df_diego['HORA_INI_STR'], format='%I:%M %p', errors='coerce').dt.time
        df_diego['HORA_FIN'] = pd.to_datetime(df_diego['HORA_FIN_STR'], format='%I:%M %p', errors='coerce').dt.time
        df_diego['FECHAS'] = pd.to_datetime(df_diego['FECHAS'], errors='coerce')
        df_diego = df_diego.drop(columns=['HORA_INI_STR', 'HORA_FIN_STR', 'HORARIO'])

        # --- DETERMINAR ESTADO DEL CURSO (Filtro inteligente) ---
        # Si tiene celdas vacías en 'ESTADO DE CLASE' -> significa que el curso está ACTIVO
        # Si todo está lleno -> el curso ha FINALIZADe
        df_estados = df_diego.groupby('ID')['ESTADO DE CLASE'].apply( #la función lambda afecta a 'ESTADO DE CLASE' 
            lambda x: 'ACTIVO' if x.isna().any() else 'FINALIZADO' # x.isna() busca celdas vacías y .any() pregunta: "¿Hay al menos una vacía?".
        ).reset_index(name='ESTADO_CURSO')
        
        df_diego = pd.merge(df_diego, df_estados, on='ID', how='left')

        # -----------------------
        # PREPARACIÓN DE VISTAS
        # -----------------------
        df_operativa = df_diego.sort_values(by=['FECHAS', 'HORA_INICIO']) # sort_values ordena las filas
        orden = ['SOPORTE', 'CURSO', 'DOCENTE', 'PERIODO', 'NRC', 'ID', 'SESIÓN', 'FECHAS', 'HORA_INICIO', 'HORA_FIN', 'ESTADO DE CLASE', 'ESTADO_CURSO']
        orden_final = [c for c in orden if c in df_operativa.columns] # filtro de seguridad por si alguna columna falla
        df_operativa = df_operativa[orden_final]

        df_resumen = df_diego.groupby(['ID', 'CURSO', 'DOCENTE', 'ESTADO_CURSO']).size().reset_index(name='Total Sesiones') # size es un contador

        # -----------------------
        # EXPORTACIÓN 1: SUPERVISAR CLASES
        # -----------------------
        print(f"   Generando '{ARCHIVO_SUPERVISAR}'...")
        # instrucción:  cada vez que encuentre una columna de tipo fecha, muestrala en excel como ...
        with pd.ExcelWriter(ARCHIVO_SUPERVISAR, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer: # with se encarga de abrir el archiv y cuando finaliza se asegura de guardad y cerrar
            df_operativa.to_excel(writer, sheet_name='operativo', index=False)
            df_resumen.to_excel(writer, sheet_name='resumen', index=False)

            workbook = writer.book
            ws_operativa = writer.sheets['operativo']
            ws_resumen = writer.sheets['resumen']
            
            f_text = workbook.add_format({'num_format': '@', 'align': 'center', 'valign': 'vcenter'})
            f_center = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

            # Formatear Operativo
            (max_f, max_c) = df_operativa.shape
            ws_operativa.add_table(0, 0, max_f, max_c - 1, {
                'columns': [{'header': c} for c in df_operativa.columns], 
                'style': 'TableStyleMedium2', 
                'name': 'TablaOperativa'
            })
            ws_operativa.set_column(0, max_c - 1, 15, f_center)
            ws_operativa.set_column('F:F', 18, f_text)
            ws_operativa.set_column('B:C', 30)

            # Formatear Resumen
            (max_fr, max_cr) = df_resumen.shape
            ws_resumen.add_table(0, 0, max_fr, max_cr - 1, {
                'columns': [{'header': c} for c in df_resumen.columns], 
                'style': 'TableStyleMedium2', 
                'name': 'TablaResumen'
            })
            ws_resumen.set_column(0, max_cr - 1, 15, f_center)
            ws_resumen.set_column('B:C', 30)
            ws_resumen.set_column('D:D', 18) 
        
        print("✅ supervisar_clases.xlsx creado correctamente.")

        # ==========================================
        # 4. FUSIÓN Y FILTRADO PARA EL BOT
        # ==========================================
        print("\n>>> Paso 3: Fusionando y Filtrando para el Bot...")
        
        if os.path.exists(FILE_MAPA_IDS):
            df_mapa = pd.read_excel(FILE_MAPA_IDS, sheet_name='Mapa', dtype={'ID': str})
            df_resumen_activos = df_resumen[df_resumen['ESTADO_CURSO'] == 'ACTIVO'].copy()
            df_resumen_activos['ID'] = df_resumen_activos['ID'].astype(str)
            
            df_final_bot = pd.merge(df_resumen_activos, df_mapa[['ID', 'ID_Interno']], on='ID', how='left')
            df_final_bot.to_excel(ARCHIVO_RESUMEN_LLAVE, index=False)
            
            finalizados_count = len(df_resumen[df_resumen['ESTADO_CURSO'] == 'FINALIZADO'])
            print(f"✅ Filtro aplicado: {len(df_resumen_activos)} cursos activos enviados al Bot.")
            print(f"ℹ️ {finalizados_count} cursos finalizados fueron omitidos para mayor velocidad.")
        else:
            print("❌ ERROR: No se encontró 'base_maestra_ids.xlsx'.")

        # ==========================================
        # 5. ALERTAS
        # ==========================================
        print("\n>>> Paso 4: Auditando Anomalías (Alertas)...")
        lista_alertas = []
        for id_val, grupo in df_diego.groupby('ID'):
            c_unicos = grupo['CURSO'].dropna().unique() # primero elimina celdas vacias, y luego extrae los valores unicos
            if len(c_unicos) > 1:
                lista_alertas.append({'ID': id_val, 'Tipo': 'Nombre Contradictorio', 'Detalle': " / ".join(str(x) for x in c_unicos), 'Acción': 'Revisar Panel'})
            d_unicos = grupo['DOCENTE'].dropna().unique()
            if len(d_unicos) > 1:
                lista_alertas.append({'ID': id_val, 'Tipo': 'Múltiples Docentes', 'Detalle': " / ".join(str(x) for x in d_unicos), 'Acción': 'Verificar reemplazo'})

        if lista_alertas:
            df_alertas = pd.DataFrame(lista_alertas)
            with pd.ExcelWriter(ARCHIVO_ALERTAS, engine='xlsxwriter') as writer:
                df_alertas.to_excel(writer, index=False, sheet_name='Alertas')
                worksheet = writer.sheets['Alertas']
                f_wrap = writer.book.add_format({'text_wrap': True, 'valign': 'top'})
                worksheet.set_column('A:A', 20, f_wrap); worksheet.set_column('B:B', 25, f_wrap); worksheet.set_column('C:C', 60, f_wrap); worksheet.set_column('D:D', 35, f_wrap)
            print(f"🚨 Alertas generadas en: {ARCHIVO_ALERTAS}")

    else:
        print("⚠️ No se encontraron registros para 'DIEGO'.")

except Exception as e:
    print(f"❌ Error Crítico: {e}")
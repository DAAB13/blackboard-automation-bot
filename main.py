import subprocess
import os
import sys

# Definimos las rutas relativas a los scripts
SCRIPT_ETL = os.path.join("scr", "01_etl_programacion.py")
SCRIPT_BOT = os.path.join("scr", "03_bot_scraper.py")

print("=================================================")
print("🚀 INICIANDO SISTEMA DE AUTOMATIZACIÓN UPN")
print("=================================================")

# ---------------------------------------------------------
# PASO 1: EJECUTAR ETL (Limpieza y Preparación)
# ---------------------------------------------------------
print("\n[1/2] 🧠 Ejecutando ETL (Limpieza de Programación)...")
try:
    # subprocess.run ejecuta el script como si lo escribieras en la terminal
    # check=True lanza un error si el script falla
    subprocess.run([sys.executable, SCRIPT_ETL], check=True)
    print("✅ ETL completado con éxito.")
except subprocess.CalledProcessError:
    print("\n❌ ERROR CRÍTICO: El proceso de ETL falló.")
    print("   El robot NO se iniciará para evitar errores.")
    input("Presiona ENTER para salir...")
    sys.exit()

# ---------------------------------------------------------
# PASO 2: EJECUTAR ROBOT (Scraping)
# ---------------------------------------------------------
print("\n[2/2] 🤖 Ejecutando Robot (Descarga de Videos)...")
try:
    subprocess.run([sys.executable, SCRIPT_BOT], check=True)
    print("\n✅ Robot finalizado con éxito.")
except subprocess.CalledProcessError:
    print("\n❌ ERROR: El Robot se detuvo inesperadamente.")
    # No salimos con exit() aquí para dejar ver el mensaje final

print("\n=================================================")
print("✨ PROCESO TOTAL FINALIZADO ✨")
print("   Revisa la carpeta '02_outputs'")
print("=================================================")
input("Presiona ENTER para cerrar esta ventana...")
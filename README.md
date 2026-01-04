# 🤖 Sistema de Automatización de Reportes - Postgrado UPN

## 📝 Descripción
Este sistema ha sido diseñado para optimizar la gestión operativa del área de Postgrado. Su función principal es eliminar la carga manual de cruzar el **Panel de Programación** con los enlaces de grabaciones de las clases en **Blackboard Collaborate**.

**Impacto:** Reduce un proceso de horas de revisión manual a un ciclo automatizado de menos de 10 minutos.

---

## 📂 Estructura del Proyecto
El sistema se organiza bajo una arquitectura modular para facilitar el mantenimiento y la integridad de los datos:

* **`00_inputs/`**: Contiene la copia del Panel de Programación obtenida de OneDrive.
* **`01_data/`**: Almacena archivos de procesamiento intermedio y agendas de supervisión.
* **`02_outputs/`**: Carpeta de destino para los productos finales y reportes de alertas.
* **`scr/`**: Contiene los scripts individuales de Python.
* **`main.py`**: Orquestador central del sistema.

---

## 🔍 Detalle Técnico y Flujo de Datos

### 1. `01_etl_programacion.py` (Procesador de Datos)
Transforma el archivo de programación desordenado en estructuras listas para el bot.

* **Archivos Necesarios (Inputs):**
    1.  `PANEL DE PROGRAMACIÓN V7.xlsx`: El script busca este maestro en OneDrive y crea una copia en `00_inputs/`.
    2.  `base_maestra_ids.xlsx`: Generado por el script 02, ubicado en `01_data/`.
* **Productos Generados (Outputs):**
    1.  **`supervisar_clases.xlsx`**: Tu agenda diaria con tablas y filtros profesionales (en `01_data/`).
    2.  **`resumen_con_llave.xlsx`**: Mapa simplificado y filtrado para que el bot trabaje a máxima velocidad (en `01_data/`).
    3.  **`reporte_alertas.xlsx`**: Aviso sobre inconsistencias de nombres o docentes (en `02_outputs/`).

### 2. `02_mapa_llave.py` (El "Cerrajero" Digital)
Obtiene las credenciales técnicas necesarias para la navegación.
* **Función:** Extrae cookies de sesión e `ID_Interno` de los cursos.
* **Producto Generado:**
    * **`base_maestra_ids.xlsx`**: Diccionario técnico guardado en `01_data/`.

### 3. `03_bot_scraper.py` (El Robot Extractor)
Navega e intercepta los enlaces de las grabaciones de manera masiva.
* **Archivos Necesarios (Inputs):**
    1.  **`resumen_con_llave.xlsx`**: Proveniente de `01_data/`. Contiene la columna clave `ID_Interno`.
* **Producto Generado:**
    * **`REPORTE_FINAL_COMPLETO.xlsx`**: Consolidado final con enlaces directos (en `02_outputs/`).

---

## 🔄 Diagrama de Proceso

```mermaid
graph TD
    %% Definición del estilo para los scripts
    classDef scriptClass fill:#E3F2FD,stroke:#1976D2,stroke-width:2px,color:#000;

    %% Nodos de archivos (Óvalos)
    A(OneDrive: Panel Programación VERSIÓN 6)
    C(base_maestra_ids.xlsx)
    D(supervisar_clases.xlsx)
    E(resumen_con_llave.xlsx)
    G(REPORTE_FINAL_COMPLETO.xlsx)
    H(reporte_lertas.xlsx)

    %% Nodos de Scripts (Rectángulos con Estilo)
    I[Script 02: mapa llave]:::scriptClass
    B[Script 01: ETL & Limpieza]:::scriptClass
    F[Script 03: bot]:::scriptClass

    %% Conexiones
    A --> B
    I --> C
    C --> B
    B --> D
    B --> E
    E --> F
    F --> G
    B --> H
import pandas as pd
import os
import subprocess
from datetime import datetime

# Archivos base
ARCHIVO_BASE = "ListadoClientesMaestro.parquet"
ARCHIVO_HISTORICO = "clientes_sunat.parquet"
ARCHIVO_RUC = "pendientes_ruc.xlsx"
ARCHIVO_DNI = "pendientes_doc.xlsx"

# Resultados de scraping
ARCHIVO_RUCS_RES = "resultado_rucs.xlsx"
ARCHIVO_DNIS_RES = "resultado_dnis.xlsx"

# Archivos de salida
fecha_hoy = datetime.now().strftime("%Y%m%d")
ARCHIVO_OK = f"clientes_scrapeados_OK_{fecha_hoy}.xlsx"
ARCHIVO_ERR = f"clientes_scrapeados_ERRORES_{fecha_hoy}.xlsx"

# Columnas fijas
COLUMNAS_FIJAS = [
    "DNI/RUC","Razón Social","Fecha Inscripción","Fecha Inicio Actividades",
    "Estado Contribuyente","Condición Contribuyente","Domicilio Fiscal",
    "Sistema Emisión","Actividad Comercio Exterior","Actividad Económica",
    "Actividad Secundaria 1","Actividad Secundaria 2","Emisor Electrónico Desde",
    "Periodo","Trabajadores","Pensionistas","Prestadores",
    "Trabajadores Estado","Estado Scraping"
]

# =============================
# FUNCIONES AUXILIARES
# =============================

def cargar_parquet(path):
    if not os.path.exists(path):
        return None
    df = pd.read_parquet(path)
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.reindex(columns=COLUMNAS_FIJAS)
    return df.astype(str)

def cargar_y_normalizar_excel(path):
    if not os.path.exists(path):
        return None
    df = pd.read_excel(path, dtype=str)
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.reindex(columns=COLUMNAS_FIJAS)
    return df.astype(str)

def ejecutar(script):
    subprocess.run(["python", script], check=True)

# =============================
# 1️⃣ CARGAR MAESTRO Y DETECTAR NUEVOS
# =============================

print("Cargando maestro...")
df_actual = cargar_parquet(ARCHIVO_BASE)
if df_actual is None:
    print("No existe archivo maestro.")
    exit()

if os.path.exists(ARCHIVO_HISTORICO):
    print("Cargando histórico...")
    df_historico = cargar_parquet(ARCHIVO_HISTORICO)
    ya_procesados = set(df_historico["DNI/RUC"].str.strip())
    df_nuevos = df_actual[~df_actual["DNI/RUC"].str.strip().isin(ya_procesados)]
else:
    print("No existe histórico, se procesará todo.")
    df_nuevos = df_actual.copy()

if df_nuevos.empty:
    print("No hay registros nuevos.")
    exit()

print(f"Nuevos registros detectados: {len(df_nuevos)}")

# =============================
# 2️⃣ SEPARAR RUC Y DNI
# =============================

df_ruc = df_nuevos[df_nuevos["DNI/RUC"].str.len() == 11].copy()
df_dni = df_nuevos[df_nuevos["DNI/RUC"].str.len() == 8].copy()

df_ruc["RUC_LIMPIO"] = df_ruc["DNI/RUC"]
df_dni["RUC_LIMPIO"] = df_dni["DNI/RUC"]

df_ruc.to_excel(ARCHIVO_RUC, index=False)
df_dni.to_excel(ARCHIVO_DNI, index=False)

print(f"RUC nuevos: {len(df_ruc)} | DNI nuevos: {len(df_dni)}")

# =============================
# 3️⃣ EJECUTAR SCRAPERS
# =============================

if len(df_ruc) > 0:
    ejecutar("SCRAPING_SUNAT_RUC.py")

if len(df_dni) > 0:
    ejecutar("SCRAPING_SUNAT_DNI.py")

# =============================
# 4️⃣ CONSOLIDAR RESULTADOS
# =============================

print("\nConsolidando resultados...")
lista_total = []

df_ruc_res = cargar_y_normalizar_excel(ARCHIVO_RUCS_RES)
if df_ruc_res is not None:
    lista_total.append(df_ruc_res)

df_dni_res = cargar_y_normalizar_excel(ARCHIVO_DNIS_RES)
if df_dni_res is not None:
    lista_total.append(df_dni_res)

if not lista_total:
    print("No se encontraron resultados de scraping.")
    exit()

df_total = pd.concat(lista_total, ignore_index=True).astype(str)

# =============================
# 5️⃣ FILTRAR SOLO OK Y SIN_DECLARACIONES
# =============================

df_total["Trabajadores Estado"] = df_total["Trabajadores Estado"].str.strip().str.upper()

df_ok = df_total[df_total["Trabajadores Estado"].isin(["OK","SIN_DECLARACIONES"])].copy()
df_err = df_total[~df_total.index.isin(df_ok.index)].copy()

print(f"Registros válidos: {len(df_ok)} | Errores: {len(df_err)}")

# =============================
# 6️⃣ ACTUALIZAR HISTÓRICO
# =============================

df_historico = cargar_parquet(ARCHIVO_HISTORICO)
if df_historico is not None:
    df_actualizado = pd.concat([df_historico, df_ok], ignore_index=True)
    df_actualizado = df_actualizado.drop_duplicates(subset=["DNI/RUC"], keep="last")
else:
    df_actualizado = df_ok.copy()

df_actualizado = df_actualizado.reindex(columns=COLUMNAS_FIJAS).astype(str)
df_actualizado.to_parquet(ARCHIVO_HISTORICO, index=False)

print(f"Histórico actualizado: {len(df_actualizado)} registros")

# =============================
# 7️⃣ EXPORTAR ARCHIVOS DEL DÍA
# =============================

df_ok.to_excel(ARCHIVO_OK, index=False)
df_err.to_excel(ARCHIVO_ERR, index=False)

print("\nPIPELINE COMPLETADO CORRECTAMENTE")
print(f"Archivo generado OK: {ARCHIVO_OK}")
print(f"Archivo generado ERRORES: {ARCHIVO_ERR}")
print(f"Nuevos agregados hoy: {len(df_ok)}")
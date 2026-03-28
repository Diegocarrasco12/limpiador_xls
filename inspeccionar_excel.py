import pandas as pd
import os

# =========================
# CONFIG
# =========================
archivo = "Planilla de seguimiento Rancagua (1).xlsx"

# =========================
# VALIDAR EXISTENCIA
# =========================
if not os.path.exists(archivo):
    print("❌ El archivo NO existe en esta carpeta")
    print("\n📂 Archivos disponibles:")
    for f in os.listdir():
        print("-", f)
    exit()

# =========================
# CARGAR EXCEL
# =========================
try:
    xls = pd.ExcelFile(archivo)
except Exception as e:
    print("❌ Error al abrir el archivo:")
    print(e)
    raise

# =========================
# MOSTRAR HOJAS
# =========================
print("\n==============================")
print("📄 HOJAS DETECTADAS")
print("==============================")
print(xls.sheet_names)

# =========================
# MOSTRAR COLUMNAS POR HOJA
# =========================
for hoja in xls.sheet_names:
    print(f"\n==============================")
    print(f"📄 HOJA: {hoja}")
    print("==============================")

    try:
        df = pd.read_excel(xls, sheet_name=hoja)

        print("\n📊 Columnas:")
        for col in df.columns:
            print("-", col)

        print("\n🔍 Primeras 5 filas:")
        print(df.head())

        print("\n📐 Tamaño (filas, columnas):")
        print(df.shape)

    except Exception as e:
        print(f"❌ Error leyendo hoja {hoja}: {e}")
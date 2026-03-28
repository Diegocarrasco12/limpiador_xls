import pandas as pd

# =========================
# CONFIG
# =========================
archivo = "Planilla de seguimiento Rancagua (1).xlsx"

# =========================
# FUNCIONES
# =========================
def limpiar_codigo(x):
    if pd.isna(x):
        return ""
    return str(x).strip().upper()

# =========================
# 1. CARGAR SOLO HOJAS UTILES
# =========================
print("📥 Cargando datos...")

df_stock = pd.read_excel(archivo, sheet_name="Stock Actual")
df_conteo_raw = pd.read_excel(archivo, sheet_name="CONTEO STOCK NOV", header=None)

# =========================
# 2. LIMPIAR STOCK ACTUAL
# =========================
df_stock["CODIGO"] = df_stock["CODIGO"].apply(limpiar_codigo)

# =========================
# 3. LIMPIAR CONTEO (FIX HEADER ROTO)
# =========================

# Buscar fila donde empieza el header real
for i, row in df_conteo_raw.iterrows():
    if "Número de artículo" in row.values:
        header_row = i
        break

# reconstruir tabla
df_conteo = pd.read_excel(
    archivo,
    sheet_name="CONTEO STOCK NOV",
    header=header_row
)

# limpiar columnas
df_conteo = df_conteo.rename(columns={
    "Número de artículo": "CODIGO",
    "total fisico": "STOCK_CONTEO"
})

df_conteo["CODIGO"] = df_conteo["CODIGO"].apply(limpiar_codigo)

# eliminar filas vacías
df_conteo = df_conteo[df_conteo["CODIGO"] != ""]

# =========================
# 4. CRUCE REAL (REEMPLAZA BUSCARV)
# =========================
print("🔗 Cruzando datos...")

df_final = df_stock.merge(
    df_conteo[["CODIGO", "STOCK_CONTEO"]],
    on="CODIGO",
    how="left"
)

# =========================
# 5. CALCULOS CORRECTOS
# =========================

df_final["DIFERENCIA"] = df_final["STOCK FÍSICO"] - df_final["Stock Disponible"]

df_final["DIF_CONTEO"] = df_final["STOCK_CONTEO"] - df_final["Stock Disponible"]

# =========================
# 6. DETECTAR ERRORES
# =========================

no_match = df_final[df_final["STOCK_CONTEO"].isna()]
duplicados = df_final[df_final.duplicated("CODIGO")]

# =========================
# 7. EXPORTAR
# =========================
print("💾 Generando archivo limpio...")

with pd.ExcelWriter("inventario_limpio.xlsx") as writer:
    df_final.to_excel(writer, sheet_name="INVENTARIO_LIMPIO", index=False)
    no_match.to_excel(writer, sheet_name="ERROR_NO_MATCH", index=False)
    duplicados.to_excel(writer, sheet_name="DUPLICADOS", index=False)

print("✅ PROCESO TERMINADO")
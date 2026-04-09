# Python
Python script 


import pandas as pd
import unicodedata


### Cargar Excel
df = pd.read_excel("VariasEmpresas_KPI´S_Todos_03_2026.xlsx", engine="openpyxl")

---
### 1️⃣ Eliminar filas con "1001" o "Totales" en columna A
df = df[~df['A'].isin(["1001", "Totales"])]

---
### 2️⃣ Reemplazar 1912 y 2281 → 24
df['A'] = df['A'].replace({"1912": "24", "2281": "24"})

---
### 3️⃣ Eliminar filas donde la columna N contiene "No cotiza S.S."
df = df[df['N'] != "No cotiza S.S."]

---
### Si columna I contiene "Baja por fusión absorción empresa"
# → vaciar columnas H e I

mask_fusion = df['I'] == "Baja por fusión absorción empresa"
df.loc[mask_fusion, ['H', 'I']] = ""

---
### 5️⃣ Si H tiene fecha y I está vacía → borrar contenido de H
# -------------------------------------------------------------------
mask_fecha_h = pd.to_datetime(df['H'], errors='coerce').notna()
mask_i_vacio = df['I'].isna() | (df['I'].astype(str).str.strip() == "")

df.loc[mask_fecha_h & mask_i_vacio, 'H'] = ""

---
### 6️⃣ Convertir formato de fecha F, G, J → YYYY-MM-DD
for col in ['F', 'G', 'J']:
    df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime("%Y-%m-%d")

---
### 7️⃣ Eliminar acentos y la "ü" en todo el DataFrame
def quitar_acentos(texto):
    if isinstance(texto, str):
        # Normalizar → eliminar tildes y diéresis
        texto = unicodedata.normalize('NFKD', texto)
        texto = texto.encode('ascii', 'ignore').decode('ascii')
    return texto

df = df.applymap(quitar_acentos)

---
### 8️⃣ En la fila 1 (cabeceras), sustituir espacios por "_"
df.columns = [col.replace(" ", "_") for col in df.columns]

---
### Guardar resultado final
df.to_excel("VariasEmpresas_KPI´S_Todos_03_2026_PROCESADO.xlsx", index=False, engine="openpyxl"


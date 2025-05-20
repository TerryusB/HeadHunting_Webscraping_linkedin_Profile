import pandas as pd

# Ruta de tu archivo
file_path = r"C:\Users\terry\OneDrive\Escritorio\Proyecto Masterclass.xlsx"

# Palabras clave
puestos = ["Gerente de Marketing", "Director de Marketing", "Jefe de Marketing", "Marketing Manager"]
sectores = ["cadena de restaurantes", "industria alimentaria", "restaurantes"]

# Cargar hojas relevantes
scrap_df = pd.read_excel(file_path, sheet_name="dataset_google-search-scrap (2)")
final_df = pd.read_excel(file_path, sheet_name="DataSet Final")

# Asegúrate de que las columnas existen y son de tipo texto
for col in ["Ultimo puesto", "Sector"]:
    if col not in final_df.columns:
        final_df[col] = ""
    final_df[col] = final_df[col].astype("object")  # Fuerza tipo texto

# Buscar y copiar coincidencias
n = min(len(scrap_df), len(final_df))  # Itera solo hasta donde ambas tablas coinciden

for i in range(n):
    row = scrap_df.iloc[i]
    text = " ".join(map(str, row.values)).lower()
    # Ultimo puesto
    for puesto in puestos:
        if puesto.lower() in text:
            final_df.at[i, "Ultimo puesto"] = puesto
            break
    # Sector
    for sector in sectores:
        if sector.lower() in text:
            final_df.at[i, "Sector"] = sector
            break

# Guardar el archivo
final_df.to_excel(r"C:\Users\terry\OneDrive\Escritorio\DataSet_Final_Actualizado.xlsx", index=False)

print("¡Listo! Revisa el archivo 'DataSet_Final_Actualizado.xlsx' en tu Escritorio.")

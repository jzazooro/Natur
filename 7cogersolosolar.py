import pandas as pd

# Cargar el archivo Excel
file_path = 'aca.xlsx'
excel_data = pd.ExcelFile(file_path)

# Cargar la hoja de datos en un DataFrame
df = excel_data.parse('Sheet1')

# Filtrar el DataFrame para incluir solo filas donde la columna "Producto" sea "S0132 - Solar"
filtered_df = df[df['Producto'] == 'S0132 - Solar']

# Guardar el DataFrame filtrado en un nuevo archivo Excel
filtered_file_path = 'filtered_aca.xlsx'
filtered_df.to_excel(filtered_file_path, index=False)

print(f"Archivo filtrado guardado en: {filtered_file_path}")

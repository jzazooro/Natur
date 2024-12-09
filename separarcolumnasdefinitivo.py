import os
import pandas as pd

# Ruta de la carpeta con los archivos Excel
folder_path = 'archivos'

# Obtener la lista de todos los archivos en la carpeta
files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

# Procesar cada archivo Excel
for file in files:
    # Ruta completa del archivo
    file_path = os.path.join(folder_path, file)

    # Cargar el archivo Excel
    excel_data = pd.ExcelFile(file_path)

    # Cargar los datos de la hoja "Detalle"
    if 'Detalle' in excel_data.sheet_names:
        df = excel_data.parse('Detalle')

        # Separar la columna "Proyecto" en dos columnas: "Número" y "Dirección"
        if 'Proyecto' in df.columns:
            split_columns = df['Proyecto'].str.split(' - ', n=1, expand=True)

            # Encontrar el índice de la columna "Proyecto"
            proyecto_index = df.columns.get_loc('Proyecto')

            # Insertar "Número" y "Dirección" justo después de "Proyecto"
            df.insert(loc=proyecto_index + 1, column='Número', value=split_columns[0])  # A la derecha de "Proyecto"
            df.insert(loc=proyecto_index + 2, column='Dirección', value=split_columns[1])  # Después de "Número"

            # Guardar el DataFrame modificado en un nuevo archivo Excel
            output_path = os.path.join(folder_path, f"ACABADO_{file}")
            df.to_excel(output_path, index=False, engine='openpyxl', float_format="%.0f")

            print(f"Archivo procesado y guardado como: {output_path}")

print("Procesamiento completado.")
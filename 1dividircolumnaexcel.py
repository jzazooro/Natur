import pandas as pd

# Ruta del archivo Excel
file_path = 'IVA_ES_S0132_122024.xlsx'

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
        df.insert(loc=proyecto_index + 1, column='Número de proyecto', value=split_columns[0])  # A la derecha de "Proyecto"
        df.insert(loc=proyecto_index + 2, column='Dirección', value=split_columns[1])  # Después de "Número"

        # Guardar el DataFrame modificado en un nuevo archivo Excel
        output_path = "ACABADO_documento.xlsx"
        df.to_excel(output_path, index=False, engine='openpyxl', float_format="%.0f")

        print(f"Archivo procesado y guardado como: {output_path}")
else:
    print("La hoja 'Detalle' no existe en el archivo proporcionado.")

print("Procesamiento completado.")

import pandas as pd

# Cargar el archivo Excel (archivo habilitado para macros)
archivo_excel = "HRNUEVO.xlsm"  # Asegúrate de colocar el nombre correcto
xls = pd.ExcelFile(archivo_excel, engine='openpyxl')

# Función para obtener valores específicos de una hoja
def obtener_valores(hoja, celdas):
    df = pd.read_excel(xls, sheet_name=hoja, header=None)  # Cargar la hoja sin encabezados
    valores = {celda: df.iloc[int(celda[1:]) - 1, ord(celda[0]) - 65] for celda in celdas}
    return valores

# Celdas a extraer
celdas_modelo = ["E2", "G4", "G5", "G55", "G58"]
celdas_tabla_balance = ["B3", "C3", "D3"]

# Extraer valores
valores_modelo = obtener_valores("MODELO", celdas_modelo)
valores_tabla_balance = obtener_valores("TABLA BALANCE", celdas_tabla_balance)

# Imprimir resultados
print("Valores de la hoja MODELO:", valores_modelo)
print("Valores de la hoja TABLA BALANCE:", valores_tabla_balance)

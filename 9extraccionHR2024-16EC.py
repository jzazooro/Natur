import pandas as pd

# Cargar el archivo Excel (archivo habilitado para macros)
archivo_excel = "HRNUEVO.xlsm"  # Reemplaza con tu archivo
xls = pd.ExcelFile(archivo_excel, engine='openpyxl')

# Función para obtener valores específicos de una hoja
def obtener_valores(hoja, celdas):
    df = pd.read_excel(xls, sheet_name=hoja, header=None)  # Cargar la hoja sin encabezados
    valores = {celda: df.iloc[int(celda[1:]) - 1, ord(celda[0]) - 65] for celda in celdas}
    return valores

# Leer la celda O1 de la hoja "MODELO"
df_modelo = pd.read_excel(xls, sheet_name="MODELO", header=None)  # Cargar hoja sin encabezados
identificador = df_modelo.iloc[0, 14]  # O1 corresponde a la fila 1 (índice 0) y la columna 15 (índice 14)
identificadorproducto = df_modelo.iloc[4, 6]  # O1 corresponde a la fila 1 (índice 0) y la columna 15 (índice 14)

# Celdas a extraer
celdas_modelo = ["E2", "G4", "G5", "G55", "G58", "G40", "G42", "G43", "G44", "G45", "G46", "G33", "G34", "G36", "O15", "O16"]
celdas_modeloPPA = ["O19", "O21"]
celdas_modeloVD = ["K3", "K5"]
celdas_tabla_balance = ["B3", "C3", "D3"]

# Verificar si la celda O1 contiene "2024-16EC"
if str(identificador).strip() == "2024-16EC":

    # Extraer valores
    valores_modelo = obtener_valores("MODELO", celdas_modelo)
    valores_tabla_balance = obtener_valores("TABLA BALANCE", celdas_tabla_balance)

    # Imprimir resultados
    print("Valores de la hoja MODELO:", valores_modelo)
    print("Valores de la hoja TABLA BALANCE:", valores_tabla_balance)

    if str(identificadorproducto).strip() == "PPA":
        valores_modeloPPA = obtener_valores("MODELO", celdas_modeloPPA)
        print("Valores de la hoja MODELO en PPA:", valores_modeloPPA)

    if str(identificadorproducto).strip() == "VD":
        valores_modeloVD = obtener_valores("MODELO", celdas_modeloVD)
        print("Valores de la hoja MODELO en VD:", valores_modeloVD)

else:
    print("La celda O1 no contiene '2024-16EC'. No es correcto el modelo de HR.")

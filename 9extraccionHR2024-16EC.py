import pandas as pd

# Cargar el archivo Excel (archivo habilitado para macros)
archivo_excel = "HRNUEVO.xlsm"  # Reemplaza con tu archivo
xls = pd.ExcelFile(archivo_excel, engine='openpyxl')

# Función para obtener valores específicos de una hoja
def obtener_valores(hoja, celdas):
    df = pd.read_excel(xls, sheet_name=hoja, header=None)  # Cargar la hoja sin encabezados
    valores = {celda: df.iloc[int(celda[1:]) - 1, ord(celda[0]) - 65] for celda in celdas}
    return valores

# Leer la celda version de la HR
df_modelo = pd.read_excel(xls, sheet_name="MODELO", header=None)  # Cargar hoja sin encabezados
identificador = df_modelo.iloc[0, 14]  # O1 corresponde a la fila 1 (índice 0) y la columna 15 (índice 14)
identificadorproducto = df_modelo.iloc[4, 6]  # O1 corresponde a la fila 1 (índice 0) y la columna 15 (índice 14)





if str(identificador).strip() == "2024-16EC" or "2024-21EC":
    # Celdas a extraer
    celdas_modelo = ["E2", "G4", "G5", "G55", "G58", "G40", "G42", "G43", "G44", "G45", "G46", "G33", "G34", "G36", "O15", "O16"]
    celdas_modeloPPA = ["O19", "O21"]
    celdas_modeloVD = ["K3", "K8"]
    celdas_hr = ["E3", "C14", "D8", "E8", "F8", "G8", "H8", "I8", "J8", "K8", "L8", "M8", "N8", "O8", "P8", "Q8", "R8", "D17", "E17", "F17", "G17", "H17", "I17", "J17", "K17", "L17", "M17", "N17", "O17", "P17", "Q17", "R17"]
    celdas_tabla_balance = ["B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15", "C16", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15", "D16", "E4", "E5", "E6", "E7", "E8", "E9", "E10", "E11", "E12", "E13", "E14", "E15", "E16"]

    # Extraer valores
    valores_modelo = obtener_valores("MODELO", celdas_modelo)
    valores_tabla_balance = obtener_valores("TABLA BALANCE", celdas_tabla_balance)

    # Imprimir resultados
    print("Valores de la hoja MODELO:", valores_modelo)
    print("Valores de la hoja TABLA BALANCE:", valores_tabla_balance)

    if str(identificadorproducto).strip() == "PPA":
        valores_modeloPPA = obtener_valores("MODELO", celdas_modeloPPA)
        print("Valores de la hoja MODELO en PPA:", valores_modeloPPA)
        valores_hr = obtener_valores("6. HR", celdas_hr)
        print("Valores de la hoja HR:", valores_hr)

    if str(identificadorproducto).strip() == "VD":
        valores_modeloVD = obtener_valores("MODELO", celdas_modeloVD)
        print("Valores de la hoja MODELO en VD:", valores_modeloVD)





if str(identificador).strip() == "2023-43EC" or "2024-05EC":
    # Celdas a extraer
    celdas_modelo = ["E2", "G4", "G5", "G55", "G40", "G42", "G43", "G44", "G45", "G46", "G33", "G34", "G36", "O15", "O16"]
    celdas_modeloPPA = ["O19", "O42"]
    celdas_modeloVD = ["K3", "K8"]
    celdas_hr = ["E3", "C14", "D8", "E8", "F8", "G8", "H8", "I8", "J8", "K8", "L8", "M8", "N8", "O8", "P8", "Q8", "R8", "D17", "E17", "F17", "G17", "H17", "I17", "J17", "K17", "L17", "M17", "N17", "O17", "P17", "Q17", "R17"]
    celdas_tabla_balance = ["B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15", "C16", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15", "D16", "E4", "E5", "E6", "E7", "E8", "E9", "E10", "E11", "E12", "E13", "E14", "E15", "E16"]

    # Extraer valores
    valores_modelo = obtener_valores("MODELO", celdas_modelo)
    valores_tabla_balance = obtener_valores("TABLA BALANCE", celdas_tabla_balance)

    # Imprimir resultados
    print("Valores de la hoja MODELO:", valores_modelo)
    print("Valores de la hoja TABLA BALANCE:", valores_tabla_balance)

    if str(identificadorproducto).strip() == "PPA":
        valores_modeloPPA = obtener_valores("MODELO", celdas_modeloPPA)
        print("Valores de la hoja MODELO en PPA:", valores_modeloPPA)
        valores_hr = obtener_valores("HR", celdas_hr)
        print("Valores de la hoja HR:", valores_hr)

    if str(identificadorproducto).strip() == "VD":
        valores_modeloVD = obtener_valores("MODELO", celdas_modeloVD)
        print("Valores de la hoja MODELO en VD:", valores_modeloVD)





if str(identificador).strip() == "2023-27EC":
    # Celdas a extraer
    celdas_modelo = ["E2", "G4", "G5", "G55", "G40", "G42", "G43", "G44", "G45", "G46", "G33", "G34", "G36", "O15", "O16"]
    celdas_modeloPPA = ["O19", "O42"]
    celdas_modeloVD = ["K3", "K5"]
    celdas_hr = ["E3", "C14", "D8", "E8", "F8", "G8", "H8", "I8", "J8", "K8", "L8", "M8", "N8", "O8", "P8", "Q8", "R8", "D17", "E17", "F17", "G17", "H17", "I17", "J17", "K17", "L17", "M17", "N17", "O17", "P17", "Q17", "R17"]
    celdas_tabla_balance = ["B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15", "C16", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15", "D16", "E4", "E5", "E6", "E7", "E8", "E9", "E10", "E11", "E12", "E13", "E14", "E15", "E16"]

    # Extraer valores
    valores_modelo = obtener_valores("MODELO", celdas_modelo)
    valores_tabla_balance = obtener_valores("TABLA BALANCE", celdas_tabla_balance)

    # Imprimir resultados
    print("Valores de la hoja MODELO:", valores_modelo)
    print("Valores de la hoja TABLA BALANCE:", valores_tabla_balance)

    if str(identificadorproducto).strip() == "PPA":
        valores_modeloPPA = obtener_valores("MODELO", celdas_modeloPPA)
        print("Valores de la hoja MODELO en PPA:", valores_modeloPPA)
        valores_hr = obtener_valores("HR", celdas_hr)
        print("Valores de la hoja HR:", valores_hr)

    if str(identificadorproducto).strip() == "VD":
        valores_modeloVD = obtener_valores("MODELO", celdas_modeloVD)
        print("Valores de la hoja MODELO en VD:", valores_modeloVD)





if str(identificador).strip() == "2023-15EC":
    # Celdas a extraer
    celdas_modelo = ["E2", "G4", "G5", "G48", "G42", "G45", "G43", "G44", "G47", "G37", "G38", "G46", "O15", "O16"]
    celdas_modeloPPA = ["O19", "O42"]
    celdas_modeloVD = ["K3", "K5"]
    celdas_hr = ["E3", "C14", "D8", "E8", "F8", "G8", "H8", "I8", "J8", "K8", "L8", "M8", "N8", "O8", "P8", "Q8", "R8", "D17", "E17", "F17", "G17", "H17", "I17", "J17", "K17", "L17", "M17", "N17", "O17", "P17", "Q17", "R17"]
    celdas_tabla_balance = ["B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15", "C16", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15", "D16", "E4", "E5", "E6", "E7", "E8", "E9", "E10", "E11", "E12", "E13", "E14", "E15", "E16"]

    # Extraer valores
    valores_modelo = obtener_valores("MODELO", celdas_modelo)
    valores_tabla_balance = obtener_valores("TABLA BALANCE", celdas_tabla_balance)

    # Imprimir resultados
    print("Valores de la hoja MODELO:", valores_modelo)
    print("Valores de la hoja TABLA BALANCE:", valores_tabla_balance)

    if str(identificadorproducto).strip() == "PPA":
        valores_modeloPPA = obtener_valores("MODELO", celdas_modeloPPA)
        print("Valores de la hoja MODELO en PPA:", valores_modeloPPA)
        valores_hr = obtener_valores("HR", celdas_hr)
        print("Valores de la hoja HR:", valores_hr)

    if str(identificadorproducto).strip() == "VD":
        valores_modeloVD = obtener_valores("MODELO", celdas_modeloVD)
        print("Valores de la hoja MODELO en VD:", valores_modeloVD)





if str(identificador).strip() == "2023-11EC":
    # Celdas a extraer
    celdas_modelo = ["E2", "G4", "G5", "G48", "G42", "G45", "G43", "G44", "G47", "G37", "G38", "G46", "O20", "O21"]
    celdas_modeloPPA = ["O24", "O47"]
    celdas_modeloVD = ["K3", "K5"]
    celdas_hr = ["E3", "C14", "D8", "E8", "F8", "G8", "H8", "I8", "J8", "K8", "L8", "M8", "N8", "O8", "P8", "Q8", "R8", "D17", "E17", "F17", "G17", "H17", "I17", "J17", "K17", "L17", "M17", "N17", "O17", "P17", "Q17", "R17"]
    celdas_tabla_balance = ["B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15", "C16", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15", "D16", "E4", "E5", "E6", "E7", "E8", "E9", "E10", "E11", "E12", "E13", "E14", "E15", "E16"]

    # Extraer valores
    valores_modelo = obtener_valores("MODELO", celdas_modelo)
    valores_tabla_balance = obtener_valores("TABLA BALANCE", celdas_tabla_balance)

    # Imprimir resultados
    print("Valores de la hoja MODELO:", valores_modelo)
    print("Valores de la hoja TABLA BALANCE:", valores_tabla_balance)

    if str(identificadorproducto).strip() == "PPA":
        valores_modeloPPA = obtener_valores("MODELO", celdas_modeloPPA)
        print("Valores de la hoja MODELO en PPA:", valores_modeloPPA)
        valores_hr = obtener_valores("HR", celdas_hr)
        print("Valores de la hoja HR:", valores_hr)

    if str(identificadorproducto).strip() == "VD":
        valores_modeloVD = obtener_valores("MODELO", celdas_modeloVD)
        print("Valores de la hoja MODELO en VD:", valores_modeloVD)





if str(identificador).strip() == "2022-45EC":
    # Celdas a extraer
    celdas_modelo = ["E2", "G4", "G5", "G47", "G41", "G44", "G42", "G34", "G46", "G36", "G37", "G45", "O20", "O21"]
    celdas_modeloPPA = ["O24", "O47"]
    celdas_modeloVD = ["K3", "K4"]
    celdas_hr = ["E3", "C14", "D8", "E8", "F8", "G8", "H8", "I8", "J8", "K8", "L8", "M8", "N8", "O8", "P8", "Q8", "R8", "D17", "E17", "F17", "G17", "H17", "I17", "J17", "K17", "L17", "M17", "N17", "O17", "P17", "Q17", "R17"]
    celdas_tabla_balance = ["B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15", "C16", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15", "D16", "E4", "E5", "E6", "E7", "E8", "E9", "E10", "E11", "E12", "E13", "E14", "E15", "E16"]

    # Extraer valores
    valores_modelo = obtener_valores("MODELO", celdas_modelo)
    valores_tabla_balance = obtener_valores("TABLA BALANCE", celdas_tabla_balance)

    # Imprimir resultados
    print("Valores de la hoja MODELO:", valores_modelo)
    print("Valores de la hoja TABLA BALANCE:", valores_tabla_balance)

    if str(identificadorproducto).strip() == "PPA":
        valores_modeloPPA = obtener_valores("MODELO", celdas_modeloPPA)
        print("Valores de la hoja MODELO en PPA:", valores_modeloPPA)
        valores_hr = obtener_valores("HR", celdas_hr)
        print("Valores de la hoja HR:", valores_hr)

    if str(identificadorproducto).strip() == "VD":
        valores_modeloVD = obtener_valores("MODELO", celdas_modeloVD)
        print("Valores de la hoja MODELO en VD:", valores_modeloVD)





if str(identificador).strip() == "2022-41EC":
    # Celdas a extraer
    celdas_modelo = ["E2", "G4", "G5", "G47", "G41", "G44", "G42", "G33", "G46", "G36", "G37", "G45", "O20", "O21"]
    celdas_modeloPPA = ["O24", "O47"]
    celdas_modeloVD = ["K3", "K4"]
    celdas_hr = ["E3", "C14", "D8", "E8", "F8", "G8", "H8", "I8", "J8", "K8", "L8", "M8", "N8", "O8", "P8", "Q8", "R8", "D17", "E17", "F17", "G17", "H17", "I17", "J17", "K17", "L17", "M17", "N17", "O17", "P17", "Q17", "R17"]
    celdas_tabla_balance = ["B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15", "C16", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15", "D16", "E4", "E5", "E6", "E7", "E8", "E9", "E10", "E11", "E12", "E13", "E14", "E15", "E16"]

    # Extraer valores
    valores_modelo = obtener_valores("MODELO", celdas_modelo)
    valores_tabla_balance = obtener_valores("TABLA BALANCE", celdas_tabla_balance)

    # Imprimir resultados
    print("Valores de la hoja MODELO:", valores_modelo)
    print("Valores de la hoja TABLA BALANCE:", valores_tabla_balance)

    if str(identificadorproducto).strip() == "PPA":
        valores_modeloPPA = obtener_valores("MODELO", celdas_modeloPPA)
        print("Valores de la hoja MODELO en PPA:", valores_modeloPPA)
        valores_hr = obtener_valores("HR", celdas_hr)
        print("Valores de la hoja HR:", valores_hr)

    if str(identificadorproducto).strip() == "VD":
        valores_modeloVD = obtener_valores("MODELO", celdas_modeloVD)
        print("Valores de la hoja MODELO en VD:", valores_modeloVD)





if str(identificador).strip() == "2022-30":
    # Celdas a extraer
    celdas_modelo = ["E2", "G4", "G5", "G46", "G40", "G43", "G41", "G33", "G45", "G35", "G36", "G44", "O20", "O21"]
    celdas_modeloPPA = ["O24", "O47"]
    celdas_modeloVD = ["K3", "K4"]
    celdas_hr = ["E3", "C14", "D8", "E8", "F8", "G8", "H8", "I8", "J8", "K8", "L8", "M8", "N8", "O8", "P8", "Q8", "R8", "D17", "E17", "F17", "G17", "H17", "I17", "J17", "K17", "L17", "M17", "N17", "O17", "P17", "Q17", "R17"]
    celdas_tabla_balance = ["B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15", "C16", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15", "D16", "E4", "E5", "E6", "E7", "E8", "E9", "E10", "E11", "E12", "E13", "E14", "E15", "E16"]

    # Extraer valores
    valores_modelo = obtener_valores("MODELO", celdas_modelo)
    valores_tabla_balance = obtener_valores("TABLA BALANCE", celdas_tabla_balance)

    # Imprimir resultados
    print("Valores de la hoja MODELO:", valores_modelo)
    print("Valores de la hoja TABLA BALANCE:", valores_tabla_balance)

    if str(identificadorproducto).strip() == "PPA":
        valores_modeloPPA = obtener_valores("MODELO", celdas_modeloPPA)
        print("Valores de la hoja MODELO en PPA:", valores_modeloPPA)
        valores_hr = obtener_valores("HR", celdas_hr)
        print("Valores de la hoja HR:", valores_hr)

    if str(identificadorproducto).strip() == "VD":
        valores_modeloVD = obtener_valores("MODELO", celdas_modeloVD)
        print("Valores de la hoja MODELO en VD:", valores_modeloVD)

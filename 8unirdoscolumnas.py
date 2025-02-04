import pandas as pd

def unir_columnas_excel(archivo_entrada, archivo_salida):
    try:
        # Leer el archivo Excel
        df = pd.read_excel(archivo_entrada)
        
        # Verificar que haya al menos dos columnas
        if df.shape[1] < 2:
            raise ValueError("El archivo debe tener al menos dos columnas.")

        # Crear una nueva columna uniendo las dos primeras con " - "
        df['Unidas'] = df.iloc[:, 0].astype(str) + " - " + df.iloc[:, 1].astype(str)

        # Guardar el resultado en un nuevo archivo Excel
        df.to_excel(archivo_salida, index=False)
        print(f"Archivo guardado con éxito en: {archivo_salida}")

    except Exception as e:
        print(f"Ocurrió un error: {e}")

# Ruta del archivo original y del archivo de salida
archivo_entrada = "LISTA PROYECTOS.xlsx"
archivo_salida = "bueno.xlsx"

# Ejecutar la función
unir_columnas_excel(archivo_entrada, archivo_salida)

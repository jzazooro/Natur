import pandas as pd
import os

# Ruta de la carpeta donde están los archivos de Excel
ruta_carpeta = 'documentos'

# Obtener la lista de archivos Excel en la carpeta
archivos_excel = [archivo for archivo in os.listdir(ruta_carpeta) if archivo.endswith('.xlsx')]

# Crear un dataframe vacío para combinar los datos
df_combinado = pd.DataFrame()

# Iterar sobre los archivos y combinarlos
for archivo in archivos_excel:
    ruta_archivo = os.path.join(ruta_carpeta, archivo)
    df = pd.read_excel(ruta_archivo)  # Leer cada archivo
    df_combinado = pd.concat([df_combinado, df], ignore_index=True)  # Combinar los datos

# Guardar el archivo combinado en un nuevo archivo Excel
ruta_salida = os.path.join(ruta_carpeta, 'archivo_combinado.xlsx')
df_combinado.to_excel(ruta_salida, index=False)

print(f'Archivo combinado guardado en: {ruta_salida}')

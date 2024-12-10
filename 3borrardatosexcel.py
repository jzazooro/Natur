#Tenemos un Excel. Queremos borrar los datos de la columna "Mes inicio periodo" si en la columna "Producto" no es "S0132 - Solar".

import pandas as pd

# Cargar el archivo Excel
ruta_archivo = 'archivo.xlsx'
df = pd.read_excel(ruta_archivo)

# Condición: si el producto no es "S0132 - Solar", borrar el dato en "Mes inicio periodo"
df.loc[df['Producto'] != 'S0132 - Solar', 'Mes inicio periodo'] = None

# Guardar el resultado en un nuevo archivo Excel
ruta_salida = 'archivo_modificadofinal.xlsx'
df.to_excel(ruta_salida, index=False)

print(f'Archivo modificado guardado en: {ruta_salida}')
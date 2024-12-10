#Tenemos un Excel. Queremos ver cuantos valores distintos tiene la columna "Producto".

import pandas as pd

# Cargar el archivo Excel
ruta_archivo = 'archivo.xlsx'
df = pd.read_excel(ruta_archivo)

# Contar los valores únicos y listarlos
valores_unicos = df['Producto'].unique()
cantidad_valores = len(valores_unicos)

print(f'La columna "Producto" contiene {cantidad_valores} valores distintos:')
for valor in valores_unicos:
    print(f'- {valor}')
# Quiero hacer una lista que sea: "000001", "000002", "000003" hasta el "000170"

# Crear la lista con comillas simples
lista_numeros = [f"'{str(i).zfill(6)}'," for i in range(1, 171)]

# Unir los elementos con un salto de línea y mostrarlos
resultado = "\n".join(lista_numeros)
print(resultado)
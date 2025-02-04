# Quiero hacer una imagen cuadrada añadiendole bordes blancos para no deformar la foto

from PIL import Image, ImageOps

# Ruta de entrada y salida
image_path = "EXPLOTACION.png"  # Cambia esta ruta por la de tu imagen PNG
output_path = "EXPLOTACION2.png"  # Cambia esta ruta por donde guardarás el archivo final

# Abrir la imagen
img = Image.open(image_path)

# Obtener dimensiones
width, height = img.size

# Calcular el nuevo tamaño cuadrado
new_size = max(width, height)

# Calcular bordes simétricos
left_right_border = (new_size - width) // 2
top_bottom_border = (new_size - height) // 2

# Crear una nueva imagen cuadrada con fondo blanco
new_img = ImageOps.expand(
    img,
    border=(left_right_border, top_bottom_border, left_right_border, top_bottom_border),
    fill=(255, 255, 255)  # Blanco puro
)

# Guardar la nueva imagen en formato PNG
new_img.save(output_path)

print("Imagen procesada y guardada en:", output_path)
print(output_path)
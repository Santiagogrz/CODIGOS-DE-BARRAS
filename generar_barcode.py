import barcode
from barcode.writer import ImageWriter
import pandas as pd
from openpyxl import load_workbook

df = pd.read_excel("ARCHIVO_BARCODE.xlsx")

def generate_barcode(barcode_type, data, filename):
    # Crea un objeto de código de barras del tipo especificado
    barcode_class = barcode.get_barcode_class(barcode_type)
    barcode_instance = barcode_class(data, writer=ImageWriter())

    # Guarda el código de barras en un archivo de imagen PNG
    filename_with_extension = f"{filename}.png"
    barcode_instance.save(filename_with_extension)

    print(f"Código de barras generado: {filename_with_extension}")

# Ejemplo de uso
barcode_type = 'ean8'  # Tipo de código de barras (puedes probar otros como 'code128' o 'code39')
filename = 'barcode'   # Nombre del archivo de imagen generado (sin extensión)

for value in df["CODIGO DE MATERIAL SAP"]:
    if isinstance(value, int) and len(str(value)) == 7:
        longitud = len(value)
        filename = f"barcode_{value}"
        generate_barcode(barcode_type, str(value),filename)
    else:
        print(f"Valor inválido en la columna 'CODIGO DE MATERIAL SAP': {value}")

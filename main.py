import xlwings as xw
import logging

# Configura el logging
logging.basicConfig(
    level = logging.INFO,
    format = "%(asctime)s %(levelname)s %(message)s",
    datefmt = "%Y-%m-%d %H:%M:%S",
    filename = "log.txt" 
)

# Abre los libros de Excel
wb1 = xw.Book('doc1.xlsx')
wb2 = xw.Book('doc2.xlsx')

# Define el color de resaltado
resaltado = (250, 150, 0)

# Recorre todas las hojas de ambos libros
for sheet1, sheet2 in zip(wb1.sheets, wb2.sheets):
    # Verifica que las hojas tengan el mismo nombre
    if sheet1.name != sheet2.name:
        logging.error(f"Las hojas no coinciden: {sheet1.name} y {sheet2.name}")
        print(f"Las hojas no coinciden: {sheet1.name} y {sheet2.name}")
        continue

    # Compara las celdas en el rango utilizado
    for celda1, celda2 in zip(sheet1.used_range, sheet2.used_range):
        if celda1.value != celda2.value:
            logging.info(f"Diferencia en la celda: {celda1.address} en la hoja: {sheet1.name}")
            print(f"Diferencia en la celda: {celda1.address} en la hoja: {sheet1.name}")
            celda1.color = resaltado
            celda2.color = resaltado

# Guarda los cambios en los libros
wb1.save()
wb2.save()

# Cierra los libros de Excel
wb1.close()
wb2.close()

print("Comparación completa.")
import xlwings as xw
import logging

# Configura el logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    filename="log.txt"
)

logging.info("Proceso iniciado")

wb1 = xw.Book('doc1.xlsx')
wb2 = xw.Book('doc2.xlsx')

wb_diff = xw.Book()

#Color de resaltado
resaltado = (250, 150, 0)

# Recorre todas las hojas de ambos libros
for index, (sheet1, sheet2) in enumerate(zip(wb1.sheets, wb2.sheets), start=1):
    #Asigna el nombre a cada hoja nueva
    diff_sheet_name = f"{sheet1.name}_Dif_{index}"

    #Crea una hoja al final del libro y accede a ella
    wb_diff.sheets.add(diff_sheet_name, after=wb_diff.sheets[-1])
    ws_diff = wb_diff.sheets[diff_sheet_name]

    # Verifica que las hojas a comparar tengan el mismo nombre
    if sheet1.name != sheet2.name:
        logging.error(f"Las hojas no coinciden: {sheet1.name} y {sheet2.name}")
        continue

    # Recorre las celdas en un rango para hacerlo más rápido
    for celda1, celda2 in zip(sheet1.used_range, sheet2.used_range):
        if celda1.value != celda2.value:
            print(f"Diferencia en valor en la celda: {celda1.address} en la hoja: {sheet1.name}")
            logging.info(f"Diferencia en valor en la celda: {celda1.address} en la hoja: {sheet1.name}")

            #Asigna 1 si hay diferencia y 0 si no
            ws_diff.range(celda1.address).value = 1
            ws_diff.range(celda1.address).color = resaltado
        
        else:
            ws_diff.range(celda1.address).value = 0

if len(wb_diff.sheets) > 0:
    wb_diff.sheets[0].delete()

#Guarda el nuevo libro
wb_diff_name = "Diferencias.xlsx"
wb_diff.save(wb_diff_name)


# Cierra los libros
wb1.close()
wb2.close()
wb_diff.close()

logging.info(f"Proceso finalizado, se ha guardado el libro {wb_diff_name}")
print(f"Se han guardado las diferencias en el libro '{wb_diff_name}'")

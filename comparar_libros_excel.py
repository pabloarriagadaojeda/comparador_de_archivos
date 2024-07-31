import xlwings as xw
import logging


def comparar(ruta_libro_1, ruta_libro_2, ruta_resultado):
    # Configura el logging
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        filename="log.txt"
    )

    logging.info("\nProceso iniciado")

    try:
        wb1 = xw.Book(ruta_libro_1)
        wb2 = xw.Book(ruta_libro_2)
    except Exception as e:
        logging.error(f"Error al abrir los libros: {e}")
        return

    wb_diff = xw.Book()

    # Color de resaltado
    resaltado = (250, 150, 0)

    # Verifica si ambos libros tienen el mismo número de hojas
    if len(wb1.sheets) != len(wb2.sheets):
        logging.error("Los libros no tienen el mismo número de hojas")
        return

    # Recorre todas las hojas de ambos libros
    for index, (sheet1, sheet2) in enumerate(zip(wb1.sheets, wb2.sheets), start=1):
        # Asigna el nombre a cada hoja nueva
        diff_sheet_name = f"Diff_{index}"

        try:
            # Crea una hoja al final del libro y accede a ella
            wb_diff.sheets.add(diff_sheet_name, after=wb_diff.sheets[-1])
            ws_diff = wb_diff.sheets[diff_sheet_name]
        except Exception as e:
            logging.error(f"Error al crear o asignar nombre a la hoja: {e}")
            continue

        # Verifica que las hojas a comparar tengan el mismo nombre
        if sheet1.name != sheet2.name:
            logging.error(f"Las hojas no coinciden: {
                          sheet1.name} y {sheet2.name}")
            continue

        # Recorre las celdas en un rango para hacerlo más rápido
        for celda1, celda2 in zip(sheet1.used_range, sheet2.used_range):
            if celda1.value != celda2.value:
                print(f"Diferencia en valor en la celda: {
                      celda1.address} en la hoja: {sheet1.name}")
                logging.info(f"Diferencia en valor en la celda: {
                             celda1.address} en la hoja: {sheet1.name}")

                try:
                    # Asigna 1 si hay diferencia y 0 si no
                    ws_diff.range(celda1.address).value = 1
                    ws_diff.range(celda1.address).color = resaltado
                except Exception as e:
                    logging.error(
                        f"Error al asignar valor o color a la celda: {e}")
            else:
                try:
                    ws_diff.range(celda1.address).value = 0
                except Exception as e:
                    logging.error(f"Error al asignar valor a la celda: {e}")

    if len(wb_diff.sheets) > 0:
        try:
            wb_diff.sheets[0].delete()
        except Exception as e:
            logging.error(f"Error al eliminar la primera hoja: {e}")

    try:
        wb_diff.save(ruta_resultado)
    except Exception as e:
        logging.error(f"Error al guardar el libro de resultados: {e}")
    finally:
        wb1.close()
        wb2.close()
        wb_diff.close()

    logging.info(f"Proceso finalizado, se ha guardado el libro {
                 ruta_resultado}")
    print(f"Se han guardado las diferencias en el libro '{ruta_resultado}'")

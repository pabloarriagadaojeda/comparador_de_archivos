import pandas as pd
import logging


def consolidar(ruta_base, ruta_salida):
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        filename="log.txt",
    )

    logging.info("\nProceso iniciado")

    sheets_exclude = ['1701', '1702', '1703', '1704', '1705', '1801', '1901',
                      '1701-Proyectos', '1702-Proy Obras Urbanización', '1703-Proyecto Redes',
                      '1704-Proyectos Obras Nuevas', '1705-Proyecto Obras Modificadas',
                      '1801-Diagnostico Obra', '1901-cartografía']

    sheet_name_codes = {
        "0101-Captación Río": "0101",
        "0102-Captación Canal": "0102",
        "0103-Captación embalse": "0103",
        "0104-Captación Mar": "0104",
        "0201-Captación Drenes": "0201",
        "0202-Captación Puntera": "0202",
        "0203-Captación Sondaje": "0203",
        "0204-Captación Norias": "0204",
        "0301-PEAP A": "0301",
        "0302-PEAP B": "0302",
        "0303-PEAP C": "0303",
        "0304-PEAP D": "0304",
        "0305-PEAP E": "0305",
        "0351-PEAS": "0351",
        "0401-Estanques Enterrados": "0401",
        "0402-Estanques Elevados": "0402",
        "0501-PTAP": "0501",
        "0502-PTAP OSMOSIS": "0502",
        "0601-Sistemas Desinfección": "0601",
        "0701-Sistemas Fluoración": "0701",
        "0801-Red Distribución": "0801",
        "0802-Red Distribución Sector": "0802",
        "0803-Red Distribución Tuberias": "0803",
        "0901-Red Recolección": "0901",
        "0902-Red Recolección Sector": "0902",
        "0903-Red Recolección Tuberías": "0903",
        "0904-Red Unitarias": "0904",
        "1001-Conexión Arranques": "1001",
        "1002-Conexión Medidores": "1002",
        "1003-Conecxion UD": "1003",
        "1101-Conducción AP": "1101",
        "1102-Conducción AP Tramo": "1102",
        "1151-Consucción AS": "1151",
        "1152-Conducción AS Tramo": "1152",
        "1201-Ptas Sistemas Tratamiento": "1201",
        "1202-Ptas Pretratamiento AS": "1202",
        "1203-Ptas tratamiento primario": "1203",
        "1204-Ptas Tto. secundario": "1204",
        "1205-Ptas Desinfección Decl": "1205",
        "1206-Ptas Lodos": "1206",
        "1207-Ptas Emisarios submarinos": "1207",
        "1208-Ptas Tramo Emisarios": "1208",
        "1209-Ptas Aforos": "1209",
        "1402-Macromedidores": "1402",
        "1403-Reductores de Presion": "1403",
        "1404-Anti golpe": "1404",
        "1405-Atraviesos": "1405",
        "1501-Terrenos Recinto": "1501",
        "1502-Recintos Obra": "1502",
        "1503-Servidumbre": "1503",
        "1601-Electrógenos": "1601",
        "1602-Subestación": "1602",
        "1603-telemetría": "1603",
        "1701-Proyectos": "1701",
        "1702-Proy Obras Urbanización": "1702",
        "1703-Proyecto Redes": "1703",
        "1704-Proyectos Obras Nuevas": "1704",
        "1705-Proyecto Obras Modificadas": "1705",
        "1801-Diagnostico Obra": "1801",
        "1901-cartografía": "1901"
    }

    archivo_base = pd.ExcelFile(ruta_base)

    sheet_names = archivo_base.sheet_names

    df_final = pd.DataFrame()

    for sheet_name in sheet_names:
        if sheet_name not in sheets_exclude:
            df = pd.read_excel(archivo_base, sheet_name=sheet_name)

            if 'RUT_EMPRESA' not in df.columns:
                print(f"{sheet_name} no contiene la columna RUT_EMPRESA")
                logging.error(
                    f"{sheet_name} no contiene la columna RUT_EMPRESA")

            else:
                if df['RUT_EMPRESA'].empty:
                    print(f"{sheet_name} no tiene información en RUT_EMPRESA")
                    logging.error(
                        f"{sheet_name} no contiene información en RUT_EMPRESA")
                else:
                    df.insert(0, 'Nombre de Hoja', sheet_name)
                    df_final = pd.concat([df_final, df], ignore_index=True)
                    logging.info(f"{sheet_name} procesada correctamente")
        else:
            print(f"{sheet_name} está en la lista de exclusión, no se considerará")

    df_final.to_excel(ruta_salida, index=False, engine='openpyxl')
    logging.info("Proceso finalizado")

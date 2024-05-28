import os
import pandas as pd
from datetime import datetime
import xmltodict

# Configuración de pandas para mostrar todas las columnas
pd.set_option('display.max_columns', 10000)

# Definir las rutas de las carpetas y los nombres de los archivos de salida
carpetas = {
    'PE04': 'C:\\Users\\fenix\\Downloads\\SOF\\PE04',
    'JUICIOS': 'C:\\Users\\fenix\\Downloads\\SOF\\JUICIOS',
    'APRENDICES': 'C:\\Users\\fenix\\Downloads\\SOF\\APRENDICES',
    'PE04-QUITARDUPLICADO':'C:\\Users\\fenix\\Downloads\\SOF\\Archivo_Final\\P04-FINAL'
}

for f in carpetas:print(f)

# Directorio de salida
directorio_salida = 'C:\\Users\\fenix\\Downloads\\SOF\\Archivo_Final'

# Crear el directorio de salida si no existe
if not os.path.exists(directorio_salida):
    os.makedirs(directorio_salida)

# Directorio específico para los archivos XML convertidos
directorio_xml_convertidos = os.path.join(directorio_salida, 'P04-FINAL')

# Crear el directorio de XML convertidos si no existe
if not os.path.exists(directorio_xml_convertidos):
    os.makedirs(directorio_xml_convertidos)

# Función para convertir archivos XML a DataFrame
def xml_to_df(xml):
    with open(xml, 'r', encoding='utf-8') as f:  # Especificar la codificación como 'utf-8'
        xml_data = f.read().replace("&", "&amp;")
    xml_dict = xmltodict.parse(xml_data)

    rows = []
    for x in xml_dict["Workbook"]["Worksheet"]["Table"]["Row"][4:]:
        new_row = []
        for old_row in x["Cell"]:
            if not old_row["Data"]:
                continue
            if "#text" not in old_row["Data"]:
                new_row.append(None)
                continue
            new_row.append(old_row["Data"]["#text"])
        rows.append(new_row)
    
    df = pd.DataFrame(rows, columns=rows[0]).iloc[1:]
    return df

# Función para convertir archivos XML a Excel usando xml_to_df
def convert_xml_to_xls(ruta_carpeta, destino_carpeta):
    files = [f for f in os.listdir(ruta_carpeta) if f.endswith('.xml')]

    for f in files:
        file_path = os.path.join(ruta_carpeta, f)
        try:
            df = xml_to_df(file_path)
            nombre_archivo = os.path.splitext(f)[0] + '.xlsx'
            ruta_guardado = os.path.join(destino_carpeta, nombre_archivo)
            df.to_excel(ruta_guardado, index=False)
            print(f"Archivo XML convertido y guardado en: {ruta_guardado}")
        except Exception as e:
            print(f"Error al convertir el archivo {file_path}: {e}")

# Convertir archivos XML de la carpeta PE04 y guardarlos en P04-FINAL
convert_xml_to_xls(carpetas['PE04'], directorio_xml_convertidos)

# Función para ordenar archivos por fecha y combinarlos en un solo archivo Excel
def ordenar_combinar_y_eliminar_duplicados(carpeta_origen, carpeta_destino):
    archivos = os.listdir(carpeta_origen)
    archivos = [os.path.join(carpeta_origen, f) for f in archivos if f.endswith('.xlsx')]

    archivos_ordenados = sorted(archivos, key=lambda x: datetime.strptime(x.split('_')[-1].split('.')[0], '%Y'), reverse=True)

    contenido_visto = set()
    dfs = []
    for archivo in archivos_ordenados:
        try:
            df = pd.read_excel(archivo)
            contenido_hash = df.to_csv(index=False)
            if contenido_hash in contenido_visto:
                print(f"Archivo duplicado encontrado y omitido: {archivo}")
                continue
            contenido_visto.add(contenido_hash)
            dfs.append(df)
            print(f"Archivo añadido a la combinación: {archivo}")
        except Exception as e:
            print(f"Error al procesar el archivo {archivo}: {e}")

    if dfs:
        combined_frame = pd.concat(dfs, ignore_index=True)
        combined_frame = combined_frame.drop_duplicates()

        ruta_destino = os.path.join(carpeta_destino, 'archivofinal_combinado.xlsx')
        combined_frame.to_excel(ruta_destino, index=False)
        print(f"Archivos combinados y guardados en: {ruta_destino}")
    else:
        print("No hay objetos para concatenar. La lista 'dfs' está vacía.")


        # Función para leer archivos Excel con diferentes encabezados y agregar una columna con el nombre del archivo
def read_excel_with_header_and_filename(file_path, header, filename):
    try:
        df = pd.read_excel(file_path, header=header)
        df.insert(0, 'identificador', os.path.splitext(filename)[0])
        return df
    except Exception as e:
        print(f"Error al leer el archivo {file_path}: {e}")
        return pd.DataFrame()

# Función para combinar archivos y eliminar filas duplicadas
def combinar_y_eliminar_duplicados(nombre_carpeta, ruta_carpeta, header):
    dataframes = []

    files = [f for f in os.listdir(ruta_carpeta) if f.endswith('.xlsx') or f.endswith('.xls')]

    for f in files:
        file_path = os.path.join(ruta_carpeta, f)
        df = read_excel_with_header_and_filename(file_path, header=header, filename=f)
        if not df.empty:
            # Mostrar el tamaño del DataFrame antes de eliminar duplicados
            print(f"Tamaño del DataFrame antes de eliminar duplicados en {f}: {df.shape}")
            dataframes.append(df)

    if dataframes:
        combined_frame = pd.concat(dataframes, ignore_index=True)
        combined_frame = combined_frame.drop_duplicates()

        # Mostrar el tamaño del DataFrame después de eliminar duplicados
        print(f"Tamaño del DataFrame después de eliminar duplicados en {nombre_carpeta}: {combined_frame.shape}")

        ruta_guardado = os.path.join(directorio_salida, f'archivofinal_{nombre_carpeta}.xlsx')
        combined_frame.to_excel(ruta_guardado, index=False)
        print(f"Archivo guardado en: {ruta_guardado}")

    else:
        print(f"No se encontraron archivos Excel en la carpeta: {ruta_carpeta}")

# Procesar carpetas JUICIOS y APRENDICES
combinar_y_eliminar_duplicados('JUICIOS', carpetas['JUICIOS'], header=12)
combinar_y_eliminar_duplicados('APRENDICES', carpetas['APRENDICES'], header=4)

# Función para procesar los archivos del PE04
def procesar_archivos_pe04(ruta_carpeta_pe04, destino_carpeta):
    files_pe04 = [f for f in os.listdir(ruta_carpeta_pe04) if f.endswith('.xlsx')]

    dfs_pe04 = []
    for f in files_pe04:
        file_path = os.path.join(ruta_carpeta_pe04, f)
        df = pd.read_excel(file_path)
        if not df.empty:
            print(f"Tamaño del DataFrame antes de combinarlo en {f}: {df.shape}")
            dfs_pe04.append(df)

    if dfs_pe04:
        combined_frame_pe04 = pd.concat(dfs_pe04, ignore_index=True)
        combined_frame_pe04 = combined_frame_pe04.drop_duplicates()
        print(f"Tamaño del DataFrame combinado del PE04 después de eliminar duplicados: {combined_frame_pe04.shape}")
        ruta_guardado_pe04 = os.path.join(destino_carpeta, 'archivofinal_PE04.xlsx')
        combined_frame_pe04.to_excel(ruta_guardado_pe04, index=False)
        print(f"Archivo guardado en: {ruta_guardado_pe04}")
    else:
        print("No se encontraron archivos Excel en la carpeta PE04")

# Procesar los archivos del PE04
procesar_archivos_pe04(carpetas['PE04-QUITARDUPLICADO'], directorio_salida)

# Procesar carpetas JUICIOS y APRENDICES
combinar_y_eliminar_duplicados('JUICIOS', carpetas['JUICIOS'], header=12)
combinar_y_eliminar_duplicados('APRENDICES', carpetas['APRENDICES'], header=4)

# Directorio para archivo combinado sin duplicados
directorio_xml_combinado = os.path.join(directorio_salida, 'P04-NOREPETIDOS')

# Crear el directorio del archivo combinado si no existe
if not os.path.exists(directorio_xml_combinado):
    os.makedirs(directorio_xml_combinado)

# Ordenar, combinar y eliminar duplicados en archivos XML convertidos y guardarlos en P04-NOREPETIDOS
ordenar_combinar_y_eliminar_duplicados(directorio_xml_convertidos, directorio_xml_combinado)
import os
import pandas as pd
import shutil

# Configuración de pandas para mostrar todas las columnas
pd.set_option('display.max_columns', 10000)

# Definir las rutas de las carpetas y los nombres de los archivos de salida
carpetas = {
    'PE04': 'C:\\Users\\fenix\\Downloads\\SOF\\PE04',
    'JUICIOS': 'C:\\Users\\fenix\\Downloads\\SOF\\JUICIOS',
    'APRENDICES': 'C:\\Users\\fenix\\Downloads\\SOF\\APRENDICES'
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

# Función para cambiar la extensión de los archivos XML a XLS
def convert_xml_extension_to_xls(ruta_carpeta, destino_carpeta):
    # Obtener la lista de archivos XML en la carpeta
    files = [f for f in os.listdir(ruta_carpeta) if f.endswith('.xml')]

    for f in files:
        file_path = os.path.join(ruta_carpeta, f)
        nombre_archivo = os.path.splitext(f)[0] + '.xls'
        ruta_guardado = os.path.join(destino_carpeta, nombre_archivo)
        shutil.copy(file_path, ruta_guardado)
        print(f"Archivo XML renombrado y guardado en: {ruta_guardado}")

# Renombrar los archivos XML de la carpeta PE04 y guardarlos en P04-FINAL
convert_xml_extension_to_xls(carpetas['PE04'], directorio_xml_convertidos)

# Función para leer archivos Excel con diferentes encabezados y agregar una columna con el nombre del archivo
def read_excel_with_header_and_filename(file_path, header, filename):
    df = pd.read_excel(file_path, header=header)
    df.insert(0, 'identificador', os.path.splitext(filename)[0])
    return df

# Iterar sobre cada carpeta y combinar archivos
for nombre_carpeta, ruta_carpeta in carpetas.items():
    # Saltar el procesamiento de archivos en la carpeta PE04 ya que solo deben ser renombrados
    if nombre_carpeta == 'PE04':
        continue

    # Inicializar una lista para almacenar los DataFrames
    dataframes = []

    # Obtener la lista de archivos en la carpeta
    files = [f for f in os.listdir(ruta_carpeta) if f.endswith('.xlsx') or f.endswith('.xls')]

    # Leer y combinar los archivos de la carpeta
    for f in files:
        file_path = os.path.join(ruta_carpeta, f)
        
        if nombre_carpeta == 'JUICIOS':
            df = read_excel_with_header_and_filename(file_path, header=12, filename=f)
        elif nombre_carpeta == 'APRENDICES':
            df = read_excel_with_header_and_filename(file_path, header=4, filename=f)
        
        if not df.empty:
            dataframes.append(df)

    if dataframes:
        # Concatenar todos los DataFrames en uno solo
        combined_frame = pd.concat(dataframes, ignore_index=True)

        # Ruta para guardar el archivo combinado
        ruta_guardado = os.path.join(directorio_salida, f'archivofinal_{nombre_carpeta}.xlsx')
        combined_frame.to_excel(ruta_guardado, index=False)

        print(f"Archivo guardado en: {ruta_guardado}")
    else:
        print(f"No se encontraron archivos Excel en la carpeta: {ruta_carpeta}")


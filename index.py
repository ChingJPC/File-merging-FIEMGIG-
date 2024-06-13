import os
import pandas as pd
from datetime import datetime
import xmltodict
from customtkinter import CTk, CTkFrame, CTkEntry, CTkLabel, CTkButton, CTkCheckBox, CTkFont, CTkImage
from tkinter import PhotoImage, filedialog, messagebox, Frame, Label, Tk, Text
from PIL import Image, ImageTk
import tkinter as tk
import sys
from tkinter import ttk
from tkinter import filedialog
import time

# Redirigir stdout a un widget Text
class StdoutRedirector(object):
    def __init__(self, text_widget):
        self.text_space = text_widget

    def write(self, string):
        self.text_space.insert(tk.END, string)
        self.text_space.see(tk.END)

# Configuración de pandas para mostrar todas las columnas
pd.set_option('display.max_columns', None)

# Función para convertir archivos XML a DataFrame
def xml_to_df(xml):
    with open(xml, 'r', encoding='utf-8') as f:
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

# Función para ordenar archivos por fecha y combinarlos en un solo archivo Excel
def ordenar_combinar_y_eliminar_duplicados(carpeta_origen, carpeta_destino):
    archivos = os.listdir(carpeta_origen)
    archivos = [os.path.join(carpeta_origen, f) for f in archivos if f.endswith('.xlsx')]

    archivos_validos = []
    for archivo in archivos:
        try:
            nombre_archivo = os.path.basename(archivo)
            fecha_str = nombre_archivo.split('_')[-1].split('.')[0]
            
            if fecha_str.isdigit():
                if len(fecha_str) == 4:  # Si es un año como 2024
                    fecha = datetime.strptime(fecha_str, '%Y')
                else:  # Si es un número de secuencia como 1, 2, 3
                    fecha = int(fecha_str)
            else:
                raise ValueError("Formato de fecha no reconocido")
            
            archivos_validos.append((fecha, archivo))
        except ValueError:
            print(f"Archivo con nombre inválido para el formato de fecha: {archivo}")

    archivos_ordenados = sorted(archivos_validos, key=lambda x: x[0], reverse=True)

    contenido_visto = set()
    dfs = []
    for _, archivo in archivos_ordenados:
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
def combinar_y_eliminar_duplicados(nombre_carpeta, ruta_carpeta, header, directorio_salida):
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

def procesar():
    try:
        carpetas = {
            'PE04': entry_pe04.get(),
            'JUICIOS': entry_juicios.get(),
            'APRENDICES': entry_aprendices.get(),
            'PE04-QUITARDUPLICADO': entry_salida.get()
        }
        directorio_salida = entry_salida.get()

        # Directorio específico para los archivos XML convertidos
        directorio_xml_convertidos = os.path.join(directorio_salida, 'P04-FINAL')

        # Crear el directorio de XML convertidos si no existe
        if not os.path.exists(directorio_xml_convertidos):
            os.makedirs(directorio_xml_convertidos)

        # Mostrar pantalla de carga
        loading_screen, progress = show_loading_screen()
        total_steps = 4
        current_step = 0

        # Convertir archivos XML de la carpeta PE04 y guardarlos en P04-FINAL
        convert_xml_to_xls(carpetas['PE04'], directorio_xml_convertidos)
        current_step += 1
        update_progress(progress, (current_step / total_steps) * 100)

        # Ordenar, combinar y eliminar duplicados en archivos XML convertidos y guardarlos en P04-NOREPETIDOS
        ordenar_combinar_y_eliminar_duplicados(directorio_xml_convertidos, carpetas['PE04-QUITARDUPLICADO'])
        current_step += 1
        update_progress(progress, (current_step / total_steps) * 100)

        # Combinar archivos en la carpeta JUICIOS y eliminar duplicados
        combinar_y_eliminar_duplicados('JUICIOS', carpetas['JUICIOS'], header=12, directorio_salida=directorio_salida)
        current_step += 1
        update_progress(progress, (current_step / total_steps) * 100)

        # Combinar archivos en la carpeta APRENDICES y eliminar duplicados
        combinar_y_eliminar_duplicados('APRENDICES', carpetas['APRENDICES'], header=4, directorio_salida=directorio_salida)
        current_step += 1
        update_progress(progress, (current_step / total_steps) * 100)

        # Procesar los archivos del PE04
        procesar_archivos_pe04(directorio_xml_convertidos, carpetas['PE04-QUITARDUPLICADO'])

        # Cerrar la pantalla de carga
        loading_screen.destroy()
        
        messagebox.showinfo("Éxito", "Procesamiento completado con éxito")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error durante el procesamiento: {e}")

# Función para mostrar la pantalla de carga
def show_loading_screen():
    loading_screen = tk.Toplevel(root)
    loading_screen.title("Processing...")
    loading_screen.geometry("300x100")

    tk.Label(loading_screen, text="Processing files... Please wait.").pack(pady=10)

    progress = ttk.Progressbar(loading_screen, orient="horizontal", length=250, mode="determinate")
    progress.pack(pady=20)

    # Simulación de progreso
    progress["maximum"] = 100
    progress["value"] = 0

    root.update_idletasks()
    
    return loading_screen, progress

# Función para actualizar la barra de progreso
def update_progress(progress, value):
    progress["value"] = value
    root.update_idletasks()

# Función para seleccionar una carpeta
def seleccionar_carpeta(tipo):
    carpeta = filedialog.askdirectory()
    if carpeta:
        if tipo == 'PE04':
            entry_pe04.delete(0, tk.END)
            entry_pe04.insert(0, carpeta)
        elif tipo == 'JUICIOS':
            entry_juicios.delete(0, tk.END)
            entry_juicios.insert(0, carpeta)
        elif tipo == 'APRENDICES':
            entry_aprendices.delete(0, tk.END)
            entry_aprendices.insert(0, carpeta)
        elif tipo == 'SALIDA':
            entry_salida.delete(0, tk.END)
            entry_salida.insert(0, carpeta)

# Crear la ventana principal
root = CTk()
root.title("File merging (FIEMG)")
root.configure(bg='black')
root.config(bg='#EFEFEF')
root.resizable(False, False)

# Iconos del Software
logogeneral = PhotoImage(file='imagenes/filemergin.png')
logo = PhotoImage(file='imagenes/senalogo.png')

# Redimensionar la imagen
new_size = (150, 150)  # Nuevo tamaño (ancho, alto)
image = Image.open('imagenes/filemergin.png')
resized_image = image.resize(new_size)
logogeneral = ImageTk.PhotoImage(resized_image)

# Crear el frame para el logo grande
frame_logo = CTkFrame(root, bg_color='#721630', fg_color='#721630')
frame_logo.pack(fill=tk.BOTH)

# Asignar la imagen al label y expandirlo para ocupar toda la celda del grid
label_logo = Label(frame_logo, image=logogeneral, bg='#721630')
label_logo.pack(side=tk.LEFT, padx=10)

# Crear el frame para los widgets de la interfaz
frame_widgets = CTkFrame(root, bg_color=root.cget('bg'), fg_color=root.cget('bg'))
frame_widgets.pack(fill=tk.BOTH, pady=(100, 0))

# Definir una fuente personalizada
fuente_subtitulos = CTkFont(family="Helvetica", size=15, weight="bold")
fuente_rutas = CTkFont(family="Helvetica", size=12, slant="italic")

# Crear widgets de la interfaz
CTkLabel(frame_widgets, text="Selecciona la carpeta de PE04: ", text_color="#721630", font=fuente_subtitulos, bg_color=root.cget('bg')).grid(row=0, column=0, padx=(20,50), pady=(20,20), sticky='E')
entry_pe04 = CTkEntry(frame_widgets, width=300, font=fuente_rutas, border_color='#721630', fg_color="#A91D3A", bg_color=root.cget('bg'))
entry_pe04.grid(row=0, column=1, padx=(0,50), pady=(20, 20), sticky='W')
CTkButton(frame_widgets, text="Seleccionar", fg_color='#721630', bg_color=root.cget('bg'), command=lambda: seleccionar_carpeta('PE04')).grid(row=0, column=2, padx=(0,20), pady=(20,20), sticky='W')

CTkLabel(frame_widgets, text="Selecciona la carpeta de JUICIOS: ", text_color="#721630", font=fuente_subtitulos, bg_color=root.cget('bg')).grid(row=1, column=0, padx=(20,50), pady=20, sticky='E')
entry_juicios = CTkEntry(frame_widgets, width=300, border_color='#721630', fg_color="#A91D3A", bg_color=root.cget('bg'))
entry_juicios.grid(row=1, column=1, padx=(0,50), pady=20, sticky='W')
CTkButton(frame_widgets, text="Seleccionar", fg_color='#721630', bg_color=root.cget('bg'), command=lambda: seleccionar_carpeta('JUICIOS')).grid(row=1, column=2, padx=(0,20), pady=20, sticky='W')

CTkLabel(frame_widgets, text="Selecciona la carpeta de APRENDICES: ", text_color="#721630", font=fuente_subtitulos, bg_color=root.cget('bg')).grid(row=2, column=0, padx=(20,50), pady=20, sticky='E')
entry_aprendices = CTkEntry(frame_widgets, width=300, border_color='#721630', fg_color="#A91D3A", bg_color=root.cget('bg'))
entry_aprendices.grid(row=2, column=1, padx=(0,50), pady=20, sticky='W')
CTkButton(frame_widgets, text="Seleccionar", fg_color='#721630', bg_color=root.cget('bg'), command=lambda: seleccionar_carpeta('APRENDICES')).grid(row=2, column=2, padx=(0,20), pady=20, sticky='W')

CTkLabel(frame_widgets, text="Selecciona la carpeta de salida: ", text_color="#721630", font=fuente_subtitulos, bg_color=root.cget('bg')).grid(row=3, column=0, padx=(20,50), pady=20, sticky='E')
entry_salida = CTkEntry(frame_widgets, width=300, border_color='#721630', fg_color="#A91D3A", bg_color=root.cget('bg'))
entry_salida.grid(row=3, column=1, padx=(0,50), pady=20, sticky='W')
CTkButton(frame_widgets, text="Seleccionar", fg_color='#721630', bg_color=root.cget('bg'), command=lambda: seleccionar_carpeta('SALIDA')).grid(row=3, column=2, padx=(0,20), pady=20, sticky='W')
CTkButton(frame_widgets, text="Procesar Archivos", fg_color='#721630', bg_color=root.cget('bg'), font=("Helvetica", 15, "bold"), command=procesar).grid(row=4, column=2, columnspan=3, pady=30, padx=(0,50), sticky='n')

# Crear un frame para mostrar la salida del proceso
frame_output = CTkFrame(root, bg_color='#721630', fg_color=root.cget('bg'))
frame_output.pack(fill=tk.BOTH, padx=10, pady=(0,10))

# Crear un widget Text dentro del frame de salida con texto en color rojo
output_text = Text(frame_output, wrap=tk.WORD,  height=10, width=50,fg='white', bg='#721630')
output_text.pack(expand=True, fill=tk.BOTH)

# Redirigir stdout al widget Text
sys.stdout = StdoutRedirector(output_text)

# Obtener el ancho y alto de la pantalla
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Calcular las coordenadas x e y para que la ventana esté centrada
x = (screen_width - 860) // 2  # 860 es el ancho de la ventana
y = (screen_height - 700) // 2  # 700 es la altura de la ventana

# Establecer la posición de la ventana en el centro de la pantalla
root.geometry("+{}+{}".format(x, y))

# Icono Software
root.call('wm','iconphoto',root._w, logo)

# Crear una barra de progreso
progreso = ttk.Progressbar(frame_widgets, orient='horizontal', length=300, mode='determinate')

# Iniciar el bucle principal de la aplicación
root.mainloop()
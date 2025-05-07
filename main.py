import os
import re
import shutil
import pandas as pd

AUDIO_EXTENSIONS = {".mp3", ".wav", ".flac", ".aac", ".ogg", ".m4a", ".wma"}

def export_folder_names_to_excel(folder_path, output_excel):
    # Obtener la lista de carpetas dentro de la ruta dada
    folder_names = [name for name in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, name))]
    
    # Crear un DataFrame de pandas con la lista de carpetas
    df = pd.DataFrame(folder_names, columns=['Folder Names'])
    
    # Exportar el DataFrame a un archivo Excel
    df.to_excel(output_excel, index=False)
    
    print(f"Los nombres de las carpetas han sido exportados a {output_excel}.")

def natural_sort_key(s):
    # Función que divide el string en partes, separando números de texto para un orden natural
    # return [int(text) if text.isdigit() else text.lower() for text in re.split('(\d+)', s)]
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

def export_file_names_to_excel(folder_path, output_excel):
    # Obtener la lista de archivos dentro de la ruta dada en orden natural
    file_names = [name for name in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, name))]
    
    # Ordenar los nombres de archivos usando la función de ordenamiento natural
    file_names.sort(key=natural_sort_key)

    # Filtrar y modificar los nombres de archivos
    modified_names = []
    original_names = []
    
    for name in file_names:
        original_names.append(name)  # Guardar el nombre original
        # Eliminar los dígitos al inicio del nombre
        modified_name = ''.join(filter(lambda x: not x.isdigit(), name)) if name[0].isdigit() else name
        modified_names.append(os.path.splitext(modified_name)[0])  # Guardar el nombre sin extensión
    
    # Crear un DataFrame de pandas con las listas de nombres
    df = pd.DataFrame({
        'Original Names': original_names,
        'Modified Names': modified_names
    })
    
    # Exportar el DataFrame a un archivo Excel
    df.to_excel(output_excel, index=False)
    
    print(f"Los nombres de los archivos han sido exportados a {output_excel}.")

# def createExcelWithWavFiles(folderPath, excelName='wavFiles.xlsx'):
#     data = []


#     # Navigate through the folder to get subfolders and .wav files
#     for subfolder in sorted(os.listdir(folderPath), key=natural_sort_key):  # Natural order for subfolders
#         subfolderPath = os.path.join(folderPath, subfolder)
#         if os.path.isdir(subfolderPath):  # Check if it is a subfolder
#             # List and sort files in the subfolder using natural order
#             files = sorted(os.listdir(subfolderPath), key=natural_sort_key)
#             for file in files:
#                 if file.lower().endswith('.wav'):  # Check if the file is a .wav
#                     # Remove leading numbers and file extension, and apply Title Case
#                     cleanFileName = re.sub(r'^\d+[_\s-]*', '', os.path.splitext(file)[0])  # Remove leading numbers
#                     cleanFileName = re.split(r'_', cleanFileName, 1)[0]  # Keep text before first underscore
#                     cleanFileName = cleanFileName.title()  # Convert to Title Case
                    
#                     data.append({'Subfolder': subfolder.title(), 'WavFile': cleanFileName})
    

#     # Create a pandas DataFrame with the data
#     df = pd.DataFrame(data, columns=['Subfolder', 'WavFile'])

#     # Save the DataFrame to an Excel file
#     df.to_excel(excelName, index=False)
def createExcelWithWavFiles(folderPath, excelName='wavFiles.xlsx'):
    data = []

    # Navigate through the folder to get subfolders and .wav files
    for subfolder in sorted(os.listdir(folderPath), key=natural_sort_key):  # Natural order for subfolders
        subfolderPath = os.path.join(folderPath, subfolder)
        if os.path.isdir(subfolderPath):  # Check if it is a subfolder
            # List and sort files in the subfolder using natural order
            files = sorted(os.listdir(subfolderPath), key=natural_sort_key)
            for file in files:
                if file.lower().endswith('.mp3'):  # Check if the file is a .wav
                    # Remove leading numbers and file extension, and apply Title Case
                    cleanFileName = re.sub(r'^\d+[_\s-]*', '', os.path.splitext(file)[0])  # Remove leading numbers
                    cleanFileName = re.split(r'_', cleanFileName, 1)[0]  # Keep text before first underscore
                    cleanFileName = cleanFileName.title()  # Convert to Title Case
                    
                    # Initialize primary and secondary artists as None
                    primary_artist, secondary_artist = subfolder.strip(), None
                    
                    # Check for 'y' or '&' to split the artists
                    if ' y ' in subfolder.lower():
                        artists = subfolder.split(' y ', 1)
                        if artists[0] == subfolder:
                            artists = subfolder.split(' Y ', 1)
                        primary_artist = artists[0].strip().title()
                        secondary_artist = artists[1].strip().title() if len(artists) > 1 else None
                    elif ' & ' in subfolder.lower():
                        artists = subfolder.split(' & ', 1)
                        primary_artist = artists[0].strip().title()
                        secondary_artist = artists[1].strip().title() if len(artists) > 1 else None
                    data.append({
                        'Subfolder': subfolder.title(),
                        'PrimaryArtist': primary_artist,
                        'SecondaryArtist': secondary_artist,
                        'WavFile': cleanFileName
                    })

    # Create a pandas DataFrame with the data
    df = pd.DataFrame(data, columns=['Subfolder', 'PrimaryArtist', 'SecondaryArtist', 'WavFile'])

    # Save the DataFrame to an Excel file
    df.to_excel(excelName, index=False)


def search_audio_files_by_folder(directorio_base, excel_name):
    datos = []
    excel_name = excel_name + '.xlsx' 
    for raiz, _, archivos in os.walk(directorio_base):
        for archivo in archivos:
            _, ext = os.path.splitext(archivo)
            if ext.lower() in AUDIO_EXTENSIONS:
                ruta_relativa = os.path.relpath(raiz, directorio_base)
                datos.append([ruta_relativa, archivo])
    # generar excel
    df = pd.DataFrame(datos, columns=["Ruta Relativa", "Archivo"])
    df.to_excel(excel_name, index=False)
    print(f"Archivo Excel generado: {excel_name}")

def copy_files_from_excel(excel_path, source_folder, destination_folder):
    # Leer el archivo Excel
    df = pd.read_excel(excel_path, header=None)
    
    # Tomar la primera columna sin importar su nombre y eliminar la extensión
    file_names = set(df.iloc[:, 0].dropna().astype(str).apply(lambda x: os.path.splitext(x)[0]))
    
    for root, _, files in os.walk(source_folder):
        for file in files:
            file_name_without_ext = os.path.splitext(file)[0]
            if file_name_without_ext in file_names:
                relative_path = os.path.relpath(root, source_folder)
                new_directory = os.path.join(destination_folder, relative_path)
                os.makedirs(new_directory, exist_ok=True)
                shutil.copy2(os.path.join(root, file), os.path.join(new_directory, file))
                print(f"Copiado: {file} -> {new_directory}")
    
    print("Proceso de copiado finalizado.")

def combine_excels_in_folder(folder_path, output_filename):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    combined_sheets = {}

    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        try:
            xls = pd.ExcelFile(file_path)
            for sheet in xls.sheet_names:
                df = xls.parse(sheet)
                if sheet in combined_sheets:
                    combined_sheets[sheet] = pd.concat([combined_sheets[sheet], df], ignore_index=True)
                else:
                    combined_sheets[sheet] = df
        except Exception as e:
            print(f"Error processing {file}: {e}")

    # Save combined Excel file
    output_path = os.path.join(folder_path, output_filename)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name, df in combined_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Combined file saved to: {output_path}")
    return output_path
option = -1 
while option != 0:
    print("Seleccione opcion: ")
    # Genera excel con solo el nombre de las carpetas
    print("1. Obtener nombres carpetas")
    # Genera excel con solo el nombre de los archivos
    print("2. Obtener nombres archivos")
    print("3. Obtener archivos y nombre de subcarpetas")
    print("4. Obtener archivos de audio y su carpeta contenedora")
    print("5. Buscar archivos desde Excel y copiarlos manteniendo estructura")
    print("6. Juntar excels reportes video de prueba")
    print("0. Salir")
    option = int(input("-> "))

    if option == 1:
        folderPath = input("ingrese la ruta de la carpeta").replace('"', '').replace("'", "") 
        export_folder_names_to_excel(folderPath, "folderNames.xlsx")

    elif option == 2:
        folderPath = input("ingrese la ruta de la carpeta: ").replace('"', '').replace("'", "") 

        export_file_names_to_excel(folderPath, "folderFiles.xlsx")
    elif option == 3:
        folderPath = input("ingrese la ruta de la carpeta principal: ").replace('"', '').replace("'", "") 
        createExcelWithWavFiles(folderPath)
    elif option == 4:
        directorio = input("Ingrese la ruta del directorio a buscar: ").replace('"', '')
        excel_name = input("Ingrese el nombre del excel generado: ")
        search_audio_files_by_folder(directorio_base=directorio, excel_name=excel_name)
    elif option == 5:
        excel_path = input("Ingrese la ruta del archivo Excel: ").replace('"', '').replace("'", "")
        source_folder = input("Ingrese la ruta de la carpeta origen: ").replace('"', '').replace("'", "")
        destination_folder = input("Ingrese la ruta de la carpeta destino: ").replace('"', '').replace("'", "")
        copy_files_from_excel(excel_path, source_folder, destination_folder)
    elif option == 6:
        source_folder = input("Ingrese la ruta de la carpeta origen: ").replace('"', '').replace("'", "")
        combine_excels_in_folder(source_folder, "final_vp_report.xlsx")
    elif option == 0:
        print("Saliendo")
    else: 
        print("invalido")
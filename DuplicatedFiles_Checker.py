import hashlib
import os

def calcular_hash(file_path):
    sha256 = hashlib.sha256()
    with open(file_path, 'rb') as file:
        # Leer el archivo en bloques para manejar archivos grandes
        for block in iter(lambda: file.read(4096), b''):
            sha256.update(block)
    return sha256.hexdigest()

def encontrar_duplicados(lista_de_archivos):
    hash_dict = {}
    duplicados = []

    for file_path in lista_de_archivos:
        # Calcular el hash de cada archivo
        file_hash = calcular_hash(file_path)

        # Obtener solo el nombre del archivo (sin la ruta)
        file_name = os.path.basename(file_path)

        # Verificar si el hash ya está en el diccionario
        if file_hash in hash_dict:
            duplicados.append((file_name, os.path.basename(hash_dict[file_hash])))
        else:
            hash_dict[file_hash] = file_path

    return duplicados

# Ejemplo de uso
directorio_pdf = os.path.join(str(os.getcwd()), 'data')
#directorio_pdf = '/ruta/a/tu/directorio/pdf'
archivos_pdf = [os.path.join(directorio_pdf, archivo) for archivo in os.listdir(directorio_pdf) if archivo.endswith('.pdf')]

duplicados = encontrar_duplicados(archivos_pdf)

if duplicados:
    print("Se encontraron duplicados:")
    for duplicado in duplicados:
        print(f"{duplicado[0]} y {duplicado[1]} son duplicados.")
else:
    print("No se encontraron duplicados en el contenido de los archivos PDF.")

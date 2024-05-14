# Dependencias necesarias: pandas - PyMuPDF
import os
import fitz
import re
import pandas as pd

def extract_keywords(text):
    # Definir keywords para buscar
    keywords = ["RESOL-2021", "RESOL-2022", "Date:", "VISTO:"]
    found_keywords = []

    for keyword in keywords:
        if re.search(keyword, text, re.IGNORECASE):
            # Extraemos el texto alrededor de la keyword
            match = re.search(keyword, text, re.IGNORECASE)
            start = max(0, match.start() - 1)  # Extraemos 1 caracter antes de la keyword
            end = match.end() + 120  # Extraemos 120 caracteres despues del keyword

            # Extreamos X número de caracteres según la keyword encontrada
            keyword_len = len(keyword)
            if keyword in ["RESOL-2021", "RESOL-2022"]:
                end = start + keyword_len + 10
            elif keyword == "Date:":
                end = start + keyword_len + 25
            elif keyword == "VISTO:":
                end = start + keyword_len + 100
            else:
                end = start + 120

            found_keywords.append((keyword, text[start:match.start()], text[start + keyword_len:end]))
    
    return found_keywords

directory_path = "./contentedorPDFs"

# Creamos una lista/array para guardar las keywords
keywords_list = []

# Iteramos sobre todos los archivos en el directorio
for file in os.listdir(directory_path):
    file_name, file_ext = os.path.splitext(file)
    if file_ext == ".pdf":
        # Abrimos el archivo PDF
        file_path = os.path.join(directory_path, file)
        with fitz.open(file_path) as pdf:
            print(f"File: {file_name}")
            keywords = []
            for page in pdf:
                page_text = page.get_text("text")
                if page_text:
                    keywords.extend(extract_keywords(page_text.replace('\n', ' ')))

            keywords_list.append(keywords)

# Creamos un DataFrame de las keywords extraidas
df = pd.DataFrame(keywords_list, columns=["NOMBRE", "DESCRIPCION", "FECHA"])

# Mostramos el DataFrame
print(df)

# Guardamos el DataFrame en un archivo Excel
df.to_excel("testeo.xlsx", index=False)

# Print un mensaje de success en la consola
print("Data guardada en testeo.xlsx")
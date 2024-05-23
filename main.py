import os
import fitz
import re
import pandas as pd

def extract_keywords(text):
    # Definimos las keywords que necesitamos
    keywords = ["RESOL-2021", "RESOL-2022", "Date:", "Art. 1."]
    found_keywords = {
        "NOMBRE": [], 
        "FECHA": [], 
        "DESCRIPCION": []
    }

    for keyword in keywords:
        if re.search(keyword, text, re.IGNORECASE):
            # Extraemos el texto lindante a la keyword
            match = re.search(keyword, text, re.IGNORECASE)
            start = max(0, match.start())
            end = match.end()

            # Extract a specific number of characters according to the found keyword
            if keyword in ["RESOL-2021", "RESOL-2022"]:
                fragmentoEncontrado = text[start:end + 4]
                found_keywords["NOMBRE"].append(fragmentoEncontrado)
            elif keyword == "Date:":
                date_pattern = r"Date:\s*(\d{1,2}/\d{1,2}/\d{4})"
                date_match = re.search(date_pattern, text[start:], re.IGNORECASE)
                if date_match:
                    fragmentoEncontrado = date_match.group(1)  # Extract only the date part
                else:
                    fragmentoEncontrado = text[start+5:end + 11]  # Removemos 'Date:'
                found_keywords["FECHA"].append(fragmentoEncontrado)
            elif keyword == "Art. 1.":
                fragmentoEncontrado = text[start:end + 300]
                found_keywords["DESCRIPCION"].append(fragmentoEncontrado)
    
    return found_keywords

directory_path = "./contenedorPDFs"

# Creamos un array para almacenar las keywords
keywords_list = []

# Iteramos sobre los PDFs dentro la carpeta contenedora
for file in os.listdir(directory_path):
    file_name, file_ext = os.path.splitext(file)
    if file_ext == ".pdf":
        # Open the PDF file
        file_path = os.path.join(directory_path, file)
        with fitz.open(file_path) as pdf:
            print(f"Leyendo archivo: {file_name}")
            file_keywords = {
                "NOMBRE": [],
                "FECHA": [], 
                "DESCRIPCION": []
            }
            for page in pdf:
                page_text = page.get_text("text")
                if page_text:
                    extracted_keywords = extract_keywords(page_text.replace('\n', ' '))
                    for key in file_keywords:
                        file_keywords[key].extend(extracted_keywords[key])

            # Append the results to the keywords_list with the filename
            max_length = max(len(file_keywords["NOMBRE"]), len(file_keywords["FECHA"]), len(file_keywords["DESCRIPCION"]))
            for i in range(max_length):
                row = [
                    file_name,  # Agregamos el nombre del archivo a la columna
                    file_keywords["NOMBRE"][i] if i < len(file_keywords["NOMBRE"]) else None,
                    file_keywords["FECHA"][i] if i < len(file_keywords["FECHA"]) else None,
                    file_keywords["DESCRIPCION"][i] if i < len(file_keywords["DESCRIPCION"]) else None
                ]
                keywords_list.append(row)

# Create a DataFrame from the extracted keywords
df = pd.DataFrame(keywords_list, columns=["ARCHIVO", "NOMBRE", "FECHA", "DESCRIPCION"])

# Display the DataFrame
print(df)

# Save the DataFrame to an Excel file
df.to_excel("testeo.xlsx", index=False)

# Print a success message to the console
print("Data almacenada correctamente en: testeo.xlsx")

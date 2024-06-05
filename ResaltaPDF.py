import fitz  # PyMuPDF

def resaltar_texto(pdf_path, cadena_a_resaltar, output_path):
    # Abrir el archivo PDF
    documento = fitz.open(pdf_path)
    
    # Iterar sobre las páginas del documento
    for num_pagina in range(len(documento)):
        pagina = documento[num_pagina]
        texto_pagina = pagina.get_text("text")
        
        # Buscar la cadena en el texto de la página
        if cadena_a_resaltar in texto_pagina:
            areas = pagina.search_for(cadena_a_resaltar)
            
            # Resaltar las áreas encontradas
            for area in areas:
                pagina.add_highlight_annot(area)
    
    # Guardar el documento con los resaltados
    documento.save(output_path)

def leer_pdf_interactivamente_y_resaltar(pdf_path, cadena_a_resaltar, output_path):
    # Abrir el archivo PDF
    documento = fitz.open(pdf_path)
    texto = []

    # Extraer texto de cada página y agregarlo a la lista 'texto'
    for pagina in documento:
        texto.extend(pagina.get_text("text").splitlines())

    # Iterar sobre las líneas de texto extraídas
    for linea in texto:
        print(linea)  # Imprimir la línea de texto actual

        # Resaltar si se encuentra la cadena
        if cadena_a_resaltar in linea:
            print(f"Cadena encontrada y resaltada: {cadena_a_resaltar}")
            resaltar_texto(pdf_path, cadena_a_resaltar, output_path)
            print(f"Documento guardado con la cadena resaltada en: {output_path}")

        # Esperar la entrada del usuario
'''      while True:
            decision = input("Presione 'S' para la siguiente cadena o 'N' para detener: ").strip().upper()
            if decision == 'S':
                break  # Salir del bucle interno para imprimir la siguiente línea
            elif decision == 'N':
                print("Lectura detenida por el usuario.")
                return  # Terminar la función
'''
if __name__ == "__main__":
    pdf_path = "C:/SP/Internacional Hispacold/HC - PROYECTOS INTERNOS/PI-ORG-00040-RW-GENERICO - HERRAMIENTA CAF TRANSLATOR/PEDIDOS BUDAPEST/4100010885 - copia.pdf"  # Reemplaza con la ruta de tu archivo PDF
    cadena_a_resaltar = "9422201"  # Reemplaza con la cadena que deseas resaltar
    output_path = "C:/SP/Internacional Hispacold/HC - PROYECTOS INTERNOS/PI-ORG-00040-RW-GENERICO - HERRAMIENTA CAF TRANSLATOR/PEDIDOS BUDAPEST/4100010885 - resaltado.pdf"  # Reemplaza con la ruta del archivo de salida
    leer_pdf_interactivamente_y_resaltar(pdf_path, cadena_a_resaltar, output_path)
    

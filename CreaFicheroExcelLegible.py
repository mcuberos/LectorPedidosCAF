import os
import re
import pandas as pd
from tkinter import messagebox
import sys
import fitz
from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename 
from pandas import ExcelWriter
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
import pyodbc
import pdfplumber
import json



dfhead = pd.DataFrame({'OrderNumber': [],
                   'Version': [],
                   'OrderDate': [],
                   'Reference':[]})

dfdetail = pd.DataFrame({'ItemNumber': [],
                    'MaterialCode': [],
                    'Description': [],
                    'MaterialGroup':[],
                    'Contract':[],
                    'Deadline':[],
                    'Quantity':[],
                    'UnitPrice':[],
                    'Discount':[],
                    'TotalValue':[],
                    'Quantity_2':[],
                    'Deadline_2':[],
                    'QuantityDelivered':[],
                    'QuantityReceived':[],
                    'QuantityPending':[],
                    'DeliveryAddress':[]})

def read_pdf_tables(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
              if len(table)>1:
              #  print(pd.DataFrame)
                df = pd.DataFrame(table[1:],columns=table[0])
                print(df)
                #nueva_fila = pd.Series([clausula[1], clausula[2], requisito[1],requisito[2],100,requisito[3],requisito[4],requisito[3],nuevo_comentario,requisito[10],requisito[11],requisito[5],requisito[12]], index=dftemp.columns)
                #dftemp = dftemp._append(nueva_fila, ignore_index=True)
                tables.append(df)
    exceltemporal="C:/SP/Internacional Hispacold/HC - PROYECTOS INTERNOS/PI-ORG-00040-RW-GENERICO - HERRAMIENTA CAF TRANSLATOR/PEDIDOS BUDAPEST/excelPrueba.xlsx"
    writer=pd.ExcelWriter(exceltemporal,engine='xlsxwriter')
    df.to_excel(writer,'Excel temp',index=False)
    writer.close()
    return tables


def leer_pdf_interactivamente(pdf_path):
    # Abrir el archivo PDF
    documento = fitz.open(pdf_path)
    texto = []

    # Extraer texto de cada página y agregarlo a la lista 'texto'
    for pagina in documento:
        texto.extend(pagina.get_text("text").splitlines())

    # Iterar sobre las líneas de texto extraídas
    for linea in texto:
        #print(linea)  # Imprimir la línea de texto actual
        if linea == "Número pedido:":
          print("ESTA ES LA LÍNEA QUE PRECEDE AL PEDIDO")
          print(linea)
          print(numpedido)
        else:
          numpedido=linea
        # Esperar la entrada del usuario
        while True:
            decision = input("Presione 'S' para la siguiente cadena o 'N' para detener: ").strip().upper()
            if decision == 'S':
                break  # Salir del bucle interno para imprimir la siguiente línea
            elif decision == 'N':
                print("Lectura detenida por el usuario.")
                return  # Terminar la función


pdf_path="C:/SP/Internacional Hispacold/HC - PROYECTOS INTERNOS/PI-ORG-00040-RW-GENERICO - HERRAMIENTA CAF TRANSLATOR/PEDIDOS BUDAPEST/4100010885 - copia.pdf"
pdf_path2="C:/SP/Internacional Hispacold/HC - PROYECTOS INTERNOS/PI-ORG-00040-RW-GENERICO - HERRAMIENTA CAF TRANSLATOR/PEDIDOS BUDAPEST/4100038691.pdf"
read_pdf_tables(pdf_path)
#leer_pdf_interactivamente(pdf_path2)




#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# Importem les llibreries necessàries
from openpyxl import load_workbook
from datetime import datetime
from bs4 import BeautifulSoup
import pandas as pd
import requests
import time

# Definim la ruta on es troba l'arxiu Excel i el carreguem
ruta_excel = 'https://github.com/iAleix/New_home-/raw/112648ab436d7d3dd7a326695b7fbe47cd13f945/New_home.xlsx'
book = load_workbook(ruta_excel)

# Definim la pestanya on bolcar la informació i la carreguem
hoja = 'SIP'
sheet = book[hoja]

# Definim la columna de l'Excel on es troben les referències dels pisos
columna_d = sheet['E']

# Fem un llistat de les referències associades a tots els pisos presents en l'Excel
referencias_excel = [cell.value for cell in columna_d if cell.value is not None]

# Fem un contador per anotar quants pisos nous s'han incorporat
total_nous = 0

# Fem un contador per anotar quants pisos existents simplement s'han actualitzat
total_actualitzats = 0

# Fem un contador per anotar quants pisos s'han vengut
total_venuts = 0

# Fem un contador per anotar quants pisos s'han reactivat
total_reactivats = 0

# Bloc de codi que actualitza l'estat de l'Excel dels pisos venuts
for fila in range(6, sheet.max_row):
    
    # Guardem la referència del pis associat a la fila que estem iterant
    referencia = sheet.cell(row=fila, column=5).value
    
    # Guardem l'estat del pis associat a la fila que estem iterant
    valor_columna_b = sheet.cell(row=fila, column=2).value

    if referencia in referencias_excel and valor_columna_b != 'Venut':

        # Si l'estat del pis és actiu però la referència apareix a la dels pisos venguts, el canviem a 'Venut'
        sheet.cell(row=fila, column=2).value = 'Venut'
        
        # Si es cumpleix la mateixa lògica anterior, també actualizem la data en que es marca com a 'Venut'
        sheet.cell(row=fila, column=6).value = datetime.today().strftime('%Y-%m-%d')
        
        # Finalment, calculem la diferència de temps entre que es va detectar el pis fins que s'ha venut
        valor_columna_c = sheet.cell(row=fila, column=3).value
        fecha_actual = datetime.today().strftime('%Y-%m-%d')
        actualizacion = days_between(str(fecha_actual), str(valor_columna_c))
        sheet.cell(row=fila, column=7).value = actualizacion
        
        # I actualitzem el contador de pisos venguts
        total_venuts += 1
        
# Guardem els canvis fets
book.save(ruta_excel)


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

# URL del archivo Excel en GitHub
url = "https://raw.githubusercontent.com/iAleix/New_home-/main/New_home.xlsx"

# Descargar el archivo Excel
response = requests.get(url)
open("New_home.xlsx", "wb").write(response.content)

# Cargar el archivo Excel con openpyxl
wb = load_workbook("New_home.xlsx")
ws = wb.active

# Encontrar la primera fila libre en la columna B
for row in range(1, ws.max_row + 2):  # +2 por seguridad, para cubrir el caso de estar en la última fila
    if ws[f'B{row}'].value is None:  # Busca la primera celda vacía en la columna B
        ws[f'B{row}'] = 'ESTO ES UNA PRUEBA'
        break

# Guardar el archivo Excel
wb.save("New_home_prueba.xlsx")
wb.close()


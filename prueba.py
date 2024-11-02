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
ruta_excel = "New_home.xlsx"
book = load_workbook(ruta_excel)

# Definim la pestanya on bolcar la informació i la carreguem
hoja = 'SIP'
sheet = book[hoja]



# Generar el nombre con la fecha actual
fecha_actual = datetime.now().strftime("_%Y%m%d")

book.save(ruta_excel)

# Guardar el archivo con el nuevo nombre
book.save(f"New_home{fecha_actual}.xlsx")


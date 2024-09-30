#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import requests
import pandas as pd

# URL del archivo Excel en GitHub
url = "https://raw.githubusercontent.com/iAleix/New_home-/main/New_home.xlsx"

# Descargar el archivo Excel
response = requests.get(url)
open("New_home.xlsx", "wb").write(response.content)

# Ahora puedes cargarlo con pandas
df = pd.read_excel("New_home.xlsx")

df.head()

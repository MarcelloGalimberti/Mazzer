# Conversione BOM da PDM a Monday.com
# env neuraplprophet conda

import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO
import matplotlib.pyplot as plt
import plotly.express as px
import warnings
import zipfile
import re
warnings.filterwarnings('ignore')
from pandas.tseries.offsets import CustomBusinessDay


# impaginazione
st.set_page_config(layout="wide")

url_immagine = 'mazzer_logo.png'#?raw=true'

col_1, col_2 = st.columns([2, 3])

with col_1:
    st.image(url_immagine, width=400)

with col_2:
    st.title('Conversione distinta base PDM per Monday.com')

####### Caricamento dati

st.header('Caricamento dati | BOM da PDM', divider='gray')

uploaded_BOM_PDM = st.file_uploader("Carica distinta base da PDM (Excel)") 
if not uploaded_BOM_PDM:
    st.stop()

df_BOM_PDM = pd.read_excel(uploaded_BOM_PDM)

st.write('Distinta base PDM caricata:')
st.dataframe(df_BOM_PDM)

# preprocessing

# colonne rilevanti
colonne_rilevanti = ['Livello','Codice','Rev','Descrizione','Lavorazione 1','Lavorazione 2',
                     'Lavorazione 3','Lavorazione 4','Codice MP','Quantità']

df_BOM_PDM = df_BOM_PDM[colonne_rilevanti]

# rimuove righe con codice vuoto o descrizine vuota

df_BOM_PDM = df_BOM_PDM.dropna(subset=['Codice', 'Descrizione'])
df_BOM_PDM['Livello'] = 1
df_BOM_PDM['Rev'].fillna(0, inplace=True)
df_BOM_PDM['Quantità'] = df_BOM_PDM['Quantità'].astype(int)
df_BOM_PDM['Codice'] = df_BOM_PDM['Codice'].astype(str).str.strip()


st.subheader('Dati per il codice padre', divider='gray')

with st.form("dati_padre"):
    st.write("Inserisci i dati del padre della distinta base:")
    codice_padre = st.text_input("Codice padre")
    descrizione_padre = st.text_input("Descrizione padre")
    rev_padre = st.text_input("Revisione padre", value="0")
    lavorazione_1_padre = st.text_input("Lavorazione 1 padre")
    lavorazione_2_padre = st.text_input("Lavorazione 2 padre")
    lavorazione_3_padre = st.text_input("Lavorazione 3 padre")
    lavorazione_4_padre = st.text_input("Lavorazione 4 padre")
    submit_button = st.form_submit_button(label="Salva dati padre")
if not submit_button:
    st.stop()


#st.write("Dati del padre della distinta base:")
#st.write(f"Codice: {codice_padre}")
#st.write(f"Descrizione: {descrizione_padre}")
#st.write(f"Revisione: {rev_padre}")
#st.write(f"Lavorazione 1: {lavorazione_1_padre}")
#st.write(f"Lavorazione 2: {lavorazione_2_padre}")
#st.write(f"Lavorazione 3: {lavorazione_3_padre}")
#st.write(f"Lavorazione 4: {lavorazione_4_padre}")

# aggiunge il padre della distinta base
df_padre = pd.DataFrame({
    'Livello': [0],
    'Codice': [codice_padre],
    'Rev': [rev_padre],
    'Descrizione': [descrizione_padre],
    'Lavorazione 1': [lavorazione_1_padre],
    'Lavorazione 2': [lavorazione_2_padre],
    'Lavorazione 3': [lavorazione_3_padre],
    'Lavorazione 4': [lavorazione_4_padre],
    'Codice MP': [None],
    'Quantità': [1]
})
# concatena il padre alla distinta base
df_BOM_Monday = pd.concat([df_padre, df_BOM_PDM], ignore_index=True)
df_BOM_Monday['Rev'] = df_BOM_Monday['Rev'].astype(int)



st.subheader('Distinta base per Monday', divider='gray')
st.dataframe(df_BOM_Monday)
# salvataggio su excel del file df_BOM_Monday

def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Foglio1')
    return output.getvalue()
# Crea il bottone per scaricare file filtrato
monday_file = to_excel_bytes(df_BOM_Monday)
st.download_button(
    label=f"📥 Scarica distinta {codice_padre} per Monday",
    data=monday_file,
    file_name=f'{codice_padre}_monday.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

st.stop()


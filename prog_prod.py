# Processo dati Monday.com per piano di produzione
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
    st.title('Programma di produzione | Mazzer')

####### Caricamento dati

st.header('Caricamento dati | file Monday.com WORK_IN_PROCESS', divider='blue')

uploaded_monday = st.file_uploader("Carica Monday") 
if not uploaded_monday:
    st.stop()

df_raw = pd.read_excel(uploaded_monday, header=2, parse_dates=True)

# preprocessing

# colonne rilevanti
colonne_rilevanti = ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'RESP AVVIO','ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'VETRI', "PRIORITA' TAGLIO",
       'CONSEGNA CLIENTE PIANIFICATA', 'LT', 'DATA ORDINE', 'ID ORDINE',
       'VERIFICA LT']

df_raw = df_raw[colonne_rilevanti]

# Conversione delle colonne in datetime (se non già fatte)
for col in ["CONSEGNA CLIENTE PIANIFICATA", "DATA ORDINE"]:
    df_raw[col] = pd.to_datetime(df_raw[col], errors="coerce")

# "Tronca" l'orario lasciando solo la data (l'orario diventa sempre 00:00:00)
for col in ["CONSEGNA CLIENTE PIANIFICATA", "DATA ORDINE"]:
    df_raw[col] = df_raw[col].dt.normalize()

mesi = {
    "gen": "Jan", "feb": "Feb", "mar": "Mar", "apr": "Apr",
    "mag": "May", "giu": "Jun", "lug": "Jul", "ago": "Aug",
    "set": "Sep", "ott": "Oct", "nov": "Nov", "dic": "Dec"
}

def converti_data(data):
    if pd.isna(data):
        return pd.NaT
    for it, en in mesi.items():
        if data.startswith(it):
            data = data.replace(it, en)
    try:
        return pd.to_datetime(data, format="%b %d, %Y", errors="coerce")
    except Exception:
        return pd.NaT

df_raw["VERIFICA LT"] = df_raw["VERIFICA LT"].apply(converti_data)

# rimuove righe vuote
df_raw = df_raw.dropna(subset=["VERIFICA LT"])

# rimuove intestazioni ripetute per cliente
df_raw = df_raw.drop_duplicates(keep=False)
df_raw = df_raw.reset_index(drop=True)

# corregge ADDETTO PROD
df_raw["ADDETTO PROD"] = df_raw["ADDETTO PROD"].fillna("NON ASSEGNATO")

# sotto LT
df_raw['Sotto LT'] = np.where(df_raw['CONSEGNA CLIENTE PIANIFICATA'] < df_raw['VERIFICA LT'],True,False)

df_monday = df_raw.copy()

# date in formato gg-mm-AA
df_monday['CONSEGNA CLIENTE PIANIFICATA'] = df_monday['CONSEGNA CLIENTE PIANIFICATA'].dt.strftime('%d-%m-%Y')
df_monday['DATA ORDINE'] = df_monday['DATA ORDINE'].dt.strftime('%d-%m-%Y')
df_monday['VERIFICA LT'] = df_monday['VERIFICA LT'].dt.strftime('%d-%m-%Y')


# salvataggio su excel del file df_raw (Monday processato)
st.header('Salvataggio file Monday processato', divider='blue')

def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Foglio1')
    return output.getvalue()

# Crea il bottone per scaricare file filtrato
monday_processato = to_excel_bytes(df_monday)
st.download_button(
    label="📥 Scarica Monday processato",
    data=monday_processato,
    file_name='monday_processato.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)



# creazione file per ogni addetto
st.header('Creazione file per ogni addetto', divider='blue')
lista_addetti = df_raw["ADDETTO PROD"].unique().tolist()


# Selezione delle colonne da includere per ogni addetto

colonne_per_addetto = {
    'BELLATO D.':['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'VETRI'],
    "NANDO B.": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'ID ORDINE'],
    "PALMACCI L.": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'ID ORDINE'],
    "CIPOLLA": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'VETRI', 'ID ORDINE'],
    "SIMONE": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'ID ORDINE'],
    "CONSUELO": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA'],
    "PALMACCI F.": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'ID ORDINE'],
    "MONTE M.": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'VETRI'],
    "TAGLIO WJ": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', "PRIORITA' TAGLIO",'ID ORDINE'],#"PRIORITA' TAGLIO",
    "LAURETTI A.": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'VETRI'],
    "GUIDO": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'ID ORDINE'],
    "BONO F.": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'VETRI'],
    "PERIATI S.": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'ID ORDINE'],
    "ATTANASIO A.": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'VETRI'],
    "BELLOMO B.": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA'],
    "ARCI P": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA'],
    "SAMUELE B.": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'VETRI'],
    "DANIELE": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'ID ORDINE'],
    "CECI L.": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'VETRI'],
    "TABACCHINO": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'VETRI'],
    "MIRABELLI A.": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'VETRI'],
    "IOVANE D.": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'ID ORDINE'],
    "LAURETTI S.": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'VETRI'],
    "NON ASSEGNATO": ['Name', 'CODICE', 'CODICE CLIENTE', 'MODELLO COMMESSA', 'DESCRIZIONE',
       'N', 'ADDETTO PROD', 'FASE DI LAVORAZIONE', 'PRESSOPIEGATURA', 'WJ',
       'LASER TUBI', 'TORN', 'FRESA', 'VETRI', "PRIORITA' TAGLIO",
       'CONSEGNA CLIENTE PIANIFICATA', 'LT', 'DATA ORDINE', 'ID ORDINE',
       'VERIFICA LT']
}
giorni_target_per_addetto = {
    'BELLATO D.':14,
    "NANDO B.": 14,
    "PALMACCI L.": 14,
    "CIPOLLA": 14,
    "SIMONE": 0,
    "CONSUELO": 14,
    "PALMACCI F.": 21,
    "MONTE M.": 14,
    "TAGLIO WJ": 28,
    "LAURETTI A.": 14,
    "GUIDO": 14,
    "BONO F.": 14,
    "PERIATI S.": 14,
    "ATTANASIO A.": 14,
    "BELLOMO B.": 14,
    "ARCI P": 14,
    "SAMUELE B.": 14,
    "DANIELE": 14,
    "CECI L.": 14,
    "TABACCHINO": 14,
    "MIRABELLI A.": 14,
    "IOVANE D.": 14,
    "LAURETTI S.": 14,
    "NON ASSEGNATO": 0
}


# gestione periodi di chiusura
periodi_chiusura = [
    {"inizio": "2025-08-11", "fine": "2025-08-24"},  # ferie estive
    {"inizio": "2025-12-22", "fine": "2026-01-07"},  # chiusura natalizia
    # aggiungi altri periodi se vuoi...
]

# Costruzione della lista delle singole date di chiusura
chiusure = []
for periodo in periodi_chiusura:
    data_inizio = pd.to_datetime(periodo["inizio"])
    data_fine = pd.to_datetime(periodo["fine"])
    giorni = pd.date_range(start=data_inizio, end=data_fine, freq="D")
    chiusure.extend(giorni)

# Convertila in pandas DatetimeIndex senza duplicati
chiusure = pd.DatetimeIndex(chiusure).unique()
custom_bday = CustomBusinessDay(holidays=chiusure)

def calcola_data_target(row):
    giorni = giorni_target_per_addetto.get(row["ADDETTO PROD"], 0)
    data_consegna = row["CONSEGNA CLIENTE PIANIFICATA"]
    if pd.isna(data_consegna):
        return pd.NaT
    return data_consegna - giorni * custom_bday


zip_buffer = BytesIO()

with zipfile.ZipFile(zip_buffer, "w") as zip_file:
    for addetto in lista_addetti:
        df_addetto = df_raw[df_raw["ADDETTO PROD"] == addetto].copy()
        giorni_target = giorni_target_per_addetto.get(addetto, 0)
        # Calcolo DATA TARGET
        df_addetto["DATA TARGET"] = df_addetto.apply(calcola_data_target, axis=1)
        #df_addetto["DATA TARGET"] = df_addetto["CONSEGNA CLIENTE PIANIFICATA"] - pd.to_timedelta(giorni_target, unit="D")
        
        # Ordina per DATA TARGET crescente
        df_addetto = df_addetto.sort_values("DATA TARGET", ascending=True)
        
        # Scegli le colonne e aggiungi DATA TARGET se non già presente
        colonne = colonne_per_addetto.get(addetto, df_raw.columns.tolist())
        if "DATA TARGET" not in colonne:
            colonne.append("DATA TARGET")
        df_addetto_sub = df_addetto[colonne]
        
        # Crea excel in memoria, formatta le colonne data
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter', datetime_format='dd-mm-yyyy') as writer: #'yyyy-mm-dd'
            df_addetto_sub.to_excel(writer, index=False, sheet_name='Foglio1')
            
            # Applica la formattazione senza ora alle colonne data
            workbook  = writer.book
            worksheet = writer.sheets['Foglio1']
            date_format = workbook.add_format({'num_format': 'dd-mm-yyyy'}) #yyyy-mm-dd
            for col_idx, col_name in enumerate(df_addetto_sub.columns):
                if col_name in ["CONSEGNA CLIENTE PIANIFICATA", "DATA ORDINE", "DATA TARGET"]:
                    worksheet.set_column(col_idx, col_idx, 15, date_format)

        excel_buffer.seek(0)
        safe_addetto = re.sub(r'[^\w\-_\. ]', '_', str(addetto))
        zip_file.writestr(f"{safe_addetto}.xlsx", excel_buffer.read())

zip_buffer.seek(0)

st.download_button(
    label="Scarica tutti gli Excel in un file ZIP",
    data=zip_buffer,
    file_name="programma_addetti.zip",
    mime="application/zip"
)


st.stop()





# Creo un buffer zip in memoria
zip_buffer = BytesIO()

with zipfile.ZipFile(zip_buffer, "w") as zip_file:
    for addetto in lista_addetti:
        # Filtra le righe per l'addetto
        df_addetto = df_raw[df_raw["ADDETTO PROD"] == addetto]
        # Buffer Excel per l'addetto
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            df_addetto.to_excel(writer, index=False, sheet_name='Foglio1')
        excel_buffer.seek(0)
        # Pulisci il nome del file Excel
        safe_addetto = re.sub(r'[^\w\-_\. ]', '_', str(addetto))
        # Aggiungi il file Excel allo zip
        zip_file.writestr(f"{safe_addetto}.xlsx", excel_buffer.read())

zip_buffer.seek(0)

# Bottone di download unico zip
st.download_button(
    label="Scarica tutti gli Excel in un file ZIP",
    data=zip_buffer,
    file_name="addetti_monday.zip",
    mime="application/zip"
)



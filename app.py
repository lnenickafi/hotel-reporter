import streamlit as st
import pandas as pd
import io
import chardet
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")

st.title("🏨 Hotelový Reportér")
st.write("Nahrajte exportní soubor a stáhněte si upravenou uzávěrku.")

uploaded_file = st.file_uploader("Soubor (Excel nebo CSV)", type=["xls", "xlsx", "csv"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()
        
        # 1. NAČTENÍ DAT
        try:
            # Pokus o standardní Excel (xlsx/xls)
            df = pd.read_excel(io.BytesIO(file_bytes))
        except:
            # Pokud je to CSV, musíme se vypořádat s divnými konci řádků
            res = chardet.detect(file_bytes)
            enc = res['encoding'] if res['encoding'] else 'cp1250'
            
            try:
                # DEKÓDOVÁNÍ A VYČIŠTĚNÍ KONCŮ ŘÁDKŮ
                text_data = file_bytes.decode(enc, errors='replace')
                # splitlines() + join sjednotí všechny druhy konců řádků na standardní \n
                clean_text = "\n".join(text_data.splitlines())
                
                df = pd.read_csv(
                    io.StringIO(clean_text), 
                    sep=None, 
                    engine='python',

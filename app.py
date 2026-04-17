import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")
st.title("🏨 Hotelový Reportér")

uploaded_file = st.file_uploader("Nahrajte export (XLS)", type=["xls"])

if uploaded_file:
    try:
        # Přečteme soubor jako surové bajty
        raw_bytes = uploaded_file.read()
        
        # 1. DEKÓDOVÁNÍ "ŠPINAVÉHO" BINÁRNÍHO SOUBORU
        # Odstraníme nulové bajty a nečitelné znaky, které staré XLS generátory dělají
        try:
            # Zkusíme dekódovat jako UTF-16 (protože XLS interně používá WideChars)
            text_data = raw_bytes.decode('utf-16', errors='ignore')
        except:
            text_data = raw_bytes.decode('cp1250', errors='ignore')

        # 2. VYČIŠTĚNÍ TEXTU
        # Odstraníme kontrolní znaky a sjednotíme řádky
        clean_text = "".join(ch for ch in text_data if ch.isprintable() or ch in "\n\r\t")
        lines = clean_text.splitlines()
        
        # 3. REKONSTRUKCE TABULKY
        # Hledáme řádek, kde začínají data (obsahuje slovo Vystaveno)
        data_rows = []
        header_found = False
        target_header = []

        for line in lines:
            # Rozdělíme řádek podle tabulátorů nebo vícenásobných mezer
            parts = [p.strip() for p in re.split(r'\t|\s{2,}', line) if p.strip()]
            if not parts: continue
            
            if "Vystaveno" in parts and not header_found:
                target_header = parts
                header_found = True
                continue
            
            if header_found and len(parts) >= 5:
                data_rows.append(parts)

        if not header_found:
            # Pokud selhalo binární čištění, zkusíme nouzový režim přes Pandas
            # ale s ignorováním chyb formátu
            try:
                df = pd.read_excel(io.BytesIO(raw_bytes), engine='xlrd', recover=True)
            except:
                st.error("Soubor je příliš poškozený. Otevřete ho v Excelu a uložte jako .xlsx")
                st.stop()
        else:
            # Vytvoříme DataFrame z očištěných řádků
            df = pd.DataFrame(data_rows)
            # Přiřadíme hlavičku (zkrátíme nebo prodloužíme podle počtu sloupců)
            df.columns = target_header[:len(df.columns)]

        # 4. FILTRACE A ZPRACOVÁNÍ (Standardní logika)
        vyst_col = next((c for c in df.columns if "Vystaveno" in c), None)
        if not vyst_col:
            st.error("Nepodařilo se identifikovat sloupec Vystaveno.")
            st.stop()

        df['dt'] = pd.to_datetime(df[vyst_col], dayfirst=True, errors='coerce')
        df = df.dropna(subset=['dt'])
        
        min_date = df['dt'].min().date()
        st_t = datetime.combine(min_date, time(10, 0, 0))
        en_t = datetime.combine(min_date + pd.Timedelta(days=1), time(12, 0, 0))
        df_f = df[(df['dt'] >= st_t) & (df['dt'] <= en_t)].copy()

        # Převod čísel (očištění od měn a textu)
        for col in df_f.columns:
            if any(x in col for x in ["Základ", "DPH", "Celkem"]):
                df_f[col] = pd.to_numeric(df_f[col].astype(str).str.replace(',', '.').str.extract(r'([-+]?\d*\.?\d+)')[0], errors='coerce').fillna(0)

        # 5. EXPORT
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_f.to_excel(writer, sheet_name='Report', index=False)
            
        st.success("✅ Soubor byl úspěšně opraven a zpracován!")
        st.download_button(label="📥 Stáhnout opravený Excel", data=output.getvalue(), file_name="opraveny_report.xlsx")

    except Exception as e:
        st.error(f"Chyba: {e}. Doporučení: Otevřete soubor v Excelu a uložte jej jako typ '.xlsx'")

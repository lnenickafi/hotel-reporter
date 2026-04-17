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
        
        # 1. NAČTENÍ DAT (Robustní čtení pro tabulátorové i CSV formáty)
        try:
            # Zkusíme nejdřív Excel
            df = pd.read_excel(io.BytesIO(file_bytes))
        except:
            # Pokud selže, načteme jako text s detekcí kódování
            res = chardet.detect(file_bytes)
            enc = res['encoding'] if res['encoding'] else 'cp1250'
            text_data = file_bytes.decode(enc, errors='replace')
            # Sjednotíme konce řádků a načteme s automatickou detekcí oddělovače (sep=None)
            clean_text = "\n".join(text_data.splitlines())
            df = pd.read_csv(io.StringIO(clean_text), sep=None, engine='python', skipinitialspace=True)

        # Vyčištění názvů sloupců (odstranění uvozovek a mezer)
        df.columns = [str(c).replace('"', '').strip() for c in df.columns]

        # 2. HLEDÁNÍ HLAVIČKY (Důležité u posunutých tabulek)
        if 'Vystaveno' not in df.columns:
            found = False
            for i in range(min(20, len(df))):
                row = [str(val).strip() for val in df.iloc[i].values]
                if 'Vystaveno' in row:
                    df.columns = row
                    df = df.iloc[i+1:].reset_index(drop=True)
                    found = True
                    break
            if not found:
                st.error("V souboru nebyl nalezen sloupec 'Vystaveno'.")
                st.stop()

        # 3. FILTRACE ČASU (10:00 prvního dne až 12:00 druhého dne)
        df['Vystaveno_dt'] = pd.to_datetime(df['Vystaveno'], dayfirst=True, errors='coerce')
        df = df.dropna(subset=['Vystaveno_dt'])
        
        min_date = df['Vystaveno_dt'].min().date()
        start_threshold = datetime.combine(min_date, time(10, 0, 0))
        end_threshold = datetime.combine(min_date + pd.Timedelta(days=1), time(12, 0, 0))
        
        df_filtered = df[(df['Vystaveno_dt'] >= start_threshold) & (df['Vystaveno_dt'] <= end_threshold)].copy()

        # 4. MAPOVÁNÍ SLOUPCŮ (Přesně podle tvého zkopírovaného vzoru)
        cols_mapping = {
            "Vystaveno": "Vystaveno",
            "Stav": "Stav",
            "Číslo": "Číslo",
            "Variabilní symbol": "Variabilní symbol",
            "Forma úhrady": "Forma úhrady",
            "Splatnost": "Splatnost",
            "Základ 0%": "Základ 0%",
            "Základ - snížená sazba 12% (15%)": "Základ 12%",
            "DPH - snížená sazba 12% (15%)": "DPH 12%",
            "Základ - základní sazba 21%": "Základ 21%",
            "DPH - základní sazba 21%": "DPH 21%",
            "Celkem bez DPH": "Celkem bez DPH",
            "Celkem s DPH": "Celkem s DPH"
        }
        
        # Vybereme dostupné sloupce
        available_cols = [c for c in cols_mapping.keys() if c in df_filtered.columns]
        df_final = df_filtered[available_cols].rename(columns=cols_mapping)

        # Převod čísel (výměna čárky za tečku, aby Python uměl počítat)
        numeric_cols = ["Základ 0%", "Základ 12%", "DPH 12%", "Základ 21%", "DPH 21%", "Celkem bez DPH", "Celkem s DPH"]
        for col in numeric_cols:
            if col in df_final.columns:
                df_final[col] = pd.to_numeric(df_final[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)

        # 5. GENEROVÁNÍ EXCELU
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet('Report')
            bold = workbook.add_format({'bold': True})
            num_format = workbook.add_format({'num_format': '#,##0.00'})
            
            # Nadpis
            d1 = start_threshold.strftime("%d.%m.")
            d2 = end_threshold.strftime("%d.%m.%Y")
            worksheet.write('B1', f"Report tržeb od {d1} 10:00 do {d2} 12:00", bold)
            
            # Data
            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)
            
            # Součty (SUMIF)
            num_rows = len(df_final)
            if num_rows > 0:
                first_row = 4
                last_row = first_row + num_rows - 1
                sum_row = last_row + 2
                
                # Předpokládáme sloupce: E = Forma úhrady, M = Celkem s DPH (v nové tabulce)
                # Musíme najít písmeno sloupce dynamicky pro SUMIF
                cols = list(df_final.columns)
                try:
                    pay_col_idx = cols.index("Forma úhrady") + 1
                    total_col_idx = cols.index("Celkem s DPH") + 1
                    
                    def col_to_letter(n):
                        return chr(64 + n)

                    P = col_to_letter(pay_col_idx)
                    T = col_to_letter(total_col_idx)

                    f_hotovost = f'=SUMIF({P}{first_row}:{P}{last_row}, "*Hotově*", {T}{first_row}:{T}{last_row})'
                    f_karta = f'=SUMIF({P}{first_row}:{P}{last_row}, "*Kartou*", {T}{first_row}:{T}{last_row})'

                    # Zápis pod tabulku (posunuto na sloupce C a D)
                    worksheet.write(sum_row, 2, "Hotovost celkem:", bold)
                    worksheet.write_formula(sum_row, 3, f_hotovost, num_format)
                    worksheet.write(sum_row + 1, 2, "Kreditní karty celkem:", bold)
                    worksheet.write_formula(sum_row + 1, 3, f_karta, num_format)
                except:
                    pass

        st.success("✅ Report byl úspěšně vygenerován!")
        st.download_button(
            label="📥 Stáhnout upravený Excel",
            data=output.getvalue(),
            file_name=f"Uzaverka_{d1}{d2}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Chyba při zpracování: {e}")

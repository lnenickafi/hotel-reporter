import streamlit as st
import pandas as pd
import io
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")

st.title("🏨 Hotelový Reportér")
st.write("Nahrajte exportní soubor a stáhněte si upravenou uzávěrku.")

uploaded_file = st.file_uploader("Vyberte soubor (Excel nebo CSV)", type=["xls", "xlsx", "csv"])

if uploaded_file:
    try:
        # POKROČILÉ NAČÍTÁNÍ DAT
        file_bytes = uploaded_file.read()
        
        try:
            # 1. Pokus o načtení jako standardní Excel (xlsx/xls)
            df = pd.read_excel(io.BytesIO(file_bytes))
        except:
            # 2. Pokus o načtení jako CSV s různými kódováními
            encodings = ['cp1250', 'utf-8', 'iso-8859-2', 'latin2']
            success = False
            for enc in encodings:
                try:
                    df = pd.read_csv(io.BytesIO(file_bytes), sep=None, engine='python', encoding=enc, skipinitialspace=True)
                    if 'Vystaveno' in str(df.columns) or any(df.iloc[:5].astype(str).apply(lambda x: x.str.contains('Vystaveno')).any()):
                        success = True
                        break
                except:
                    continue
            
            if not success:
                st.error("Nepodařilo se rozpoznat formát souboru. Zkuste soubor v Excelu uložit jako 'Sešit Excel (.xlsx)' a nahrát znovu.")
                st.stop()

        # VYČIŠTĚNÍ DAT
        # Oříznutí mezer z názvů sloupců
        df.columns = [str(c).strip() for c in df.columns]
        
        # Hledání řádku s hlavičkou (pokud jsou nad tabulkou prázdné řádky nebo název "Prodejky")
        if 'Vystaveno' not in df.columns:
            found_header = False
            for i in range(min(15, len(df))):
                row_values = [str(x).strip() for x in df.iloc[i].values]
                if 'Vystaveno' in row_values:
                    df.columns = row_values
                    df = df.iloc[i+1:].reset_index(drop=True)
                    found_header = True
                    break
            if not found_header:
                st.error("V souboru nebyl nalezen sloupec 'Vystaveno'.")
                st.stop()

        # PŘEVOD A FILTRACE
        df['Vystaveno_dt'] = pd.to_datetime(df['Vystaveno'], dayfirst=True, errors='coerce')
        df = df.dropna(subset=['Vystaveno_dt'])

        # Časové okno: 10:00 (Den 1) - 12:00 (Den 2)
        min_date = df['Vystaveno_dt'].min().date()
        start_threshold = datetime.combine(min_date, time(10, 0, 0))
        end_threshold = datetime.combine(min_date + pd.Timedelta(days=1), time(12, 0, 0))
        
        df_filtered = df[(df['Vystaveno_dt'] >= start_threshold) & (df['Vystaveno_dt'] <= end_threshold)].copy()

        # Výběr sloupců
        cols_mapping = {
            "Vystaveno": "Vystaveno", "Stav": "Stav", "Číslo": "Číslo",
            "Variabilní symbol": "Variabilní symbol", "Forma úhrady": "Forma úhrady",
            "Splatnost": "Splatnost", "Základ 0%": "Základ 0%",
            "DPH - snížená sazba 12% (15%)": "DPH - 12%",
            "DPH - základní sazba 21%": "DPH 21%",
            "Celkem bez DPH": "Celkem bez DPH", "Celkem s DPH": "Celkem s DPH"
        }
        
        available_cols = [c for c in cols_mapping.keys() if c in df_filtered.columns]
        df_final = df_filtered[available_cols].rename(columns=cols_mapping)

        # Číselné formáty
        for col in ["Základ 0%", "DPH - 12%", "DPH 21%", "Celkem bez DPH", "Celkem s DPH"]:
            if col in df_final.columns:
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

        # GENEROVÁNÍ EXCELU
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet('Report')
            bold = workbook.add_format({'bold': True})
            num_fmt = workbook.add_format({'num_format': '#,##0.00'})
            
            d1_str = start_threshold.strftime("%d.%m.")
            d2_str = end_threshold.strftime("%d.%m.%Y")
            worksheet.write('B1', f"Tržba ze dne {d1_str} až {d2_str}", bold)
            
            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)
            
            num_rows = len(df_final)
            if num_rows > 0:
                f_row, l_row = 4, 4 + num_rows - 1
                s_row = l_row + 2
                worksheet.write(s_row, 2, "Hotovost:", bold)
                worksheet.write_formula(s_row, 3, f'=SUMIF(E{f_row}:E{l_row}, "*Hotově*", K{f_row}:K{l_row})', num_fmt)
                worksheet.write(s_row + 1, 2, "Kreditní kartou:", bold)
                worksheet.write_formula(s_row + 1, 3, f'=SUMIF(E{f_row}:E{l_row}, "*Kartou*", K{f_row}:K{l_row})', num_fmt)

        st.success("Report je hotový!")
        st.download_button(
            label="📥 Stáhnout upravený Excel",
            data=output.getvalue(),
            file_name=f"Trzba_{min_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Chyba při zpracování: {e}")

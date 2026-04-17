import streamlit as st
import pandas as pd
import io
import chardet
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")
st.title("🏨 Hotelový Reportér")
st.write("Nahrajte exportní soubor a stáhněte si upravenou uzávěrku.")

uploaded_file = st.file_uploader("Soubor (XLS, XLSX nebo CSV)", type=["xls", "xlsx", "csv"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()
        
        # 1. NAČTENÍ DAT
        try:
            # Pokus o čistý Excel
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None)
        except:
            # Pokud selže, načteme jako text s agresivním ošetřením
            res = chardet.detect(file_bytes)
            enc = res['encoding'] if res['encoding'] else 'cp1250'
            text_data = file_bytes.decode(enc, errors='replace')
            lines = text_data.splitlines()
            
            # Najdeme řádek s Vystaveno a určíme oddělovač
            h_idx = -1
            sep = ','
            for i, line in enumerate(lines):
                if "Vystaveno" in line:
                    h_idx = i
                    if line.count('\t') > line.count(',') and line.count('\t') > line.count(';'):
                        sep = '\t'
                    elif line.count(';') > line.count(','):
                        sep = ';'
                    break
            
            if h_idx == -1:
                st.error("V souboru nebyl nalezen sloupec 'Vystaveno'.")
                st.stop()
                
            # Načtení od hlavičky dolů, ignorujeme rozbité řádky v patičce
            csv_part = "\n".join(lines[h_idx:])
            df_raw = pd.read_csv(
                io.StringIO(csv_part),
                sep=sep,
                engine='python',
                on_bad_lines='skip', # TADY SE OPRAVUJE CHYBA "line 59"
                skipinitialspace=True
            )

        # Vyčištění názvů sloupců
        df_raw.columns = [str(c).replace('"', '').strip() for c in df_raw.columns]

        # 2. FILTRACE ČASU (10:00 - 12:00)
        # Najdeme sloupec s datem (může se jmenovat "Vystaveno" nebo mít uvozovky)
        date_col = next((c for c in df_raw.columns if "Vystaveno" in c), None)
        if not date_col:
            st.error("Nepodařilo se identifikovat sloupec s datem.")
            st.stop()

        df_raw['dt'] = pd.to_datetime(df_raw[date_col], dayfirst=True, errors='coerce')
        df_raw = df_raw.dropna(subset=['dt'])
        
        min_d = df_raw['dt'].min().date()
        st_t = datetime.combine(min_d, time(10, 0, 0))
        en_t = datetime.combine(min_d + pd.Timedelta(days=1), time(12, 0, 0))
        df_f = df_raw[(df_raw['dt'] >= st_t) & (df_raw['dt'] <= en_t)].copy()

        # 3. CHYTRÉ MAPOVÁNÍ SLOUPCŮ (hledáme shodu v názvu)
        target_cols = {
            "Vystaveno": "Vystaveno",
            "Stav": "Stav",
            "Číslo": "Číslo",
            "Variabilní symbol": "Variabilní symbol",
            "Forma úhrady": "Forma úhrady",
            "Splatnost": "Splatnost",
            "Základ 0%": "Základ 0%",
            "12%": "DPH - 12%",
            "21%": "DPH 21%",
            "Celkem bez DPH": "Celkem bez DPH",
            "Celkem s DPH": "Celkem s DPH"
        }
        
        final_map = {}
        for key, final_name in target_cols.items():
            for real_col in df_f.columns:
                if key.lower() in real_col.lower():
                    final_map[real_col] = final_name
                    break
        
        df_final = df_f[list(final_map.keys())].rename(columns=final_map)

        # Převod na čísla
        for col in df_final.columns:
            if any(x in col for x in ["Základ", "DPH", "Celkem"]):
                df_final[col] = pd.to_numeric(df_final[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)

        # 4. EXPORT DO EXCELU
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb, ws = writer.book, writer.book.add_worksheet('Report')
            f_bold, f_num = wb.add_format({'bold': True}), wb.add_format({'num_format': '#,##0.00'})
            
            d1_s, d2_s = st_t.strftime("%d.%m."), en_t.strftime("%d.%m.%Y")
            ws.write('B1', f"Tržba {d1_s} (10:00) až {d2_s} (12:00)", f_bold)
            
            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)
            
            n_rows = len(df_final)
            if n_rows > 0:
                fr, lr, sr = 4, 4 + n_rows - 1, 4 + n_rows + 1
                # SUMIF: E = Forma úhrady, K = Celkem s DPH
                ws.write(sr, 2, "Hotovost:", f_bold)
                ws.write_formula(sr, 3, f'=SUMIF(E{fr}:E{lr}, "*Hotově*", K{fr}:K{lr})', f_num)
                ws.write(sr + 1, 2, "Kreditní kartou:", f_bold)
                ws.write_formula(sr + 1, 3, f'=SUMIF(E{fr}:E{lr}, "*Kartou*", K{fr}:K{lr})', f_num)

        st.success("✅ Report úspěšně vytvořen!")
        st.download_button(
            label="📥 Stáhnout Excel",
            data=output.getvalue(),
            file_name=f"Report_{d1_s}{d2_s}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Chyba při zpracování: {e}")

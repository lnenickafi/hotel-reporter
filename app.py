import streamlit as st
import pandas as pd
import io
import chardet
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")
st.title("🏨 Hotelový Reportér")

uploaded_file = st.file_uploader("Nahrajte soubor", type=["xls", "xlsx", "csv"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()
        df = None

        # 1. NAČTENÍ - Zkusíme nejdříve Excel (odolnější proti new-line chybě)
        try:
            # Přečteme vše bez hlavičky, abychom ji našli sami
            df = pd.read_excel(io.BytesIO(file_bytes), header=None)
        except:
            # Pokud to není Excel, zkusíme CSV s ošetřením konců řádků
            res = chardet.detect(file_bytes)
            enc = res['encoding'] if res['encoding'] else 'cp1250'
            text = file_bytes.decode(enc, errors='replace').splitlines()
            # Ručně sjednotíme řádky do tabulky
            data = [line.split(',') if ',' in line else line.split(';') for line in text]
            df = pd.DataFrame(data)

        # 2. HLEDÁNÍ HLAVIČKY V CELÉ TABULCE
        h_idx = None
        for i in range(len(df)):
            row_vals = [str(v).strip() for v in df.iloc[i].values]
            if any("Vystaveno" in v for v in row_vals):
                h_idx = i
                df.columns = row_vals
                break
        
        if h_idx is None:
            st.error("Sloupec 'Vystaveno' nebyl nalezen.")
            st.stop()
            
        # Oříznutí tabulky od hlavičky dolů
        df = df.iloc[h_idx + 1:].reset_index(drop=True)
        # Vyčištění názvů sloupců
        df.columns = [str(c).strip() for c in df.columns]

        # 3. FILTRACE ČASU (10:00 - 12:00 druhý den)
        # Najdeme přesný název sloupce Vystaveno (může tam být bordel)
        vystaveno_col = next(c for c in df.columns if "Vystaveno" in c)
        df['dt'] = pd.to_datetime(df[vystaveno_col], dayfirst=True, errors='coerce')
        df = df.dropna(subset=['dt'])
        
        min_date = df['dt'].min().date()
        st_t = datetime.combine(min_date, time(10, 0, 0))
        en_t = datetime.combine(min_date + pd.Timedelta(days=1), time(12, 0, 0))
        df_f = df[(df['dt'] >= st_t) & (df['dt'] <= en_t)].copy()

        # 4. VÝBĚR SLOUPCŮ (hledáme shodu v názvu)
        mapping = {
            "Vystaveno": "Vystaveno",
            "Stav": "Stav",
            "Číslo": "Číslo",
            "Var. symbol": "Variabilní symbol",
            "Forma úhrady": "Forma úhrady",
            "DUZP": "Splatnost",
            "Základ 0%": "Základ 0%",
            "12%": "DPH - 12%",
            "21%": "DPH 21%",
            "Celkem bez DPH": "Celkem bez DPH",
            "Celkem s DPH": "Celkem s DPH"
        }
        
        final_cols = []
        for src, target in mapping.items():
            for real_col in df_f.columns:
                if src.lower() in real_col.lower():
                    df_f = df_f.rename(columns={real_col: target})
                    final_cols.append(target)
                    break
        
        df_final = df_f[final_cols].copy()

        # Převod na čísla
        for col in df_final.columns:
            if any(x in col for x in ["Základ", "DPH", "Celkem"]):
                df_final[col] = pd.to_numeric(df_final[col].astype(str).str.replace(',', '.').str.extract('(\d+\.?\d*)')[0], errors='coerce').fillna(0)

        # 5. EXPORT
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb, ws = writer.book, writer.book.add_worksheet('Report')
            f_bold, f_num = wb.add_format({'bold': True}), wb.add_format({'num_format': '#,##0.00'})
            
            d1, d2 = st_t.strftime("%d.%m."), en_t.strftime("%d.%m.%Y")
            ws.write('B1', f"Tržba {d1} - {d2}", f_bold)
            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)
            
            n_rows = len(df_final)
            if n_rows > 0:
                fr, lr, sr = 4, 4 + n_rows - 1, 4 + n_rows + 1
                # SUMIF - Forma úhrady (sloupec E), Celkem s DPH (sloupec K)
                ws.write(sr, 2, "Hotovost:", f_bold)
                ws.write_formula(sr, 3, f'=SUMIF(E{fr}:E{lr}, "*Hotově*", K{fr}:K{lr})', f_num)
                ws.write(sr + 1, 2, "Karty:", f_bold)
                ws.write_formula(sr + 1, 3, f'=SUMIF(E{fr}:E{lr}, "*Kartou*", K{fr}:K{lr})', f_num)

        st.success("Report připraven!")
        st.download_button(label="📥 Stáhnout Excel", data=output.getvalue(), file_name=f"Report_{d1}{d2}.xlsx")
        
    except Exception as e:
        st.error(f"Chyba: {e}")

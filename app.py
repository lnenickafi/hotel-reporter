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
        
        # 1. Načtení dat
        try:
            df = pd.read_excel(io.BytesIO(file_bytes))
        except:
            res = chardet.detect(file_bytes)
            enc = res['encoding'] if res['encoding'] else 'cp1250'
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), sep=None, engine='python', encoding=enc, skipinitialspace=True)
            except:
                df = pd.read_csv(io.BytesIO(file_bytes), sep=None, engine='python', encoding='cp1250', skipinitialspace=True)

        df.columns = [str(c).strip() for c in df.columns]

        # 2. Hledání hlavičky
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

        # 3. Filtrace (10:00 - 12:00)
        df['Vystaveno_dt'] = pd.to_datetime(df['Vystaveno'], dayfirst=True, errors='coerce')
        df = df.dropna(subset=['Vystaveno_dt'])
        min_d = df['Vystaveno_dt'].min().date()
        st_t = datetime.combine(min_d, time(10, 0, 0))
        en_t = datetime.combine(min_d + pd.Timedelta(days=1), time(12, 0, 0))
        df_f = df[(df['Vystaveno_dt'] >= st_t) & (df['Vystaveno_dt'] <= en_t)].copy()

        # 4. Sloupce
        mapuj = {
            "Vystaveno": "Vystaveno", "Stav": "Stav", "Číslo": "Číslo",
            "Variabilní symbol": "Variabilní symbol", "Forma úhrady": "Forma úhrady",
            "Splatnost": "Splatnost", "Základ 0%": "Základ 0%",
            "DPH - snížená sazba 12% (15%)": "DPH - 12%",
            "DPH - základní sazba 21%": "DPH 21%",
            "Celkem bez DPH": "Celkem bez DPH", "Celkem s DPH": "Celkem s DPH"
        }
        avail = [c for c in mapuj.keys() if c in df_f.columns]
        df_final = df_f[avail].rename(columns=mapuj)

        for col in ["Základ 0%", "DPH - 12%", "DPH 21%", "Celkem bez DPH", "Celkem s DPH"]:
            if col in df_final.columns:
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

        # 5. Excel export
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb = writer.book
            ws = wb.add_worksheet('Report')
            f_bold = wb.add_format({'bold': True})
            f_num = wb.add_format({'num_format': '#,##0.00'})
            
            d1, d2 = st_t.strftime("%d.%m."), en_t.strftime("%d.%m.%Y")
            ws.write('B1', f"Tržba ze dne {d1} až {d2}", f_bold)
            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)
            
            n_rows = len(df_final)
            if n_rows > 0:
                fr, lr = 4, 4 + n_rows - 1
                sr = lr + 2
                # Vzorce SUMIF
                f_hot = f'=SUMIF(E{fr}:E{lr}, "*Hotově*", K{fr}:K{lr})'
                f_kar = f'=SUMIF(E{fr}:E{lr}, "*Kartou*", K{fr}:K{lr})'
                
                ws.write(sr, 2, "Hotovost:", f_bold)
                ws.write_formula(sr, 3, f_hot, f_num)
                ws.write(sr + 1, 2, "Kreditní kartou:", f_bold)
                ws.write_formula(sr + 1, 3, f_kar, f_num)

        st.success("Hotovo!")
        st.download_button(
            label="📥 Stáhnout Excel",
            data=output.getvalue(),
            file_name=f"Report_{d1}{d2}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Chyba: {e}")

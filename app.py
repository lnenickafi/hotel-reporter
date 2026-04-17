import streamlit as st
import pandas as pd
import io
import chardet
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")
st.title("🏨 Hotelový Reportér")
st.write("Nahrajte raw exportní soubor.")

uploaded_file = st.file_uploader("Soubor (XLS, XLSX nebo CSV)", type=["xls", "xlsx", "csv"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()
        
        # 1. NAČTENÍ DAT
        try:
            # Zkusíme Excel
            df = pd.read_excel(io.BytesIO(file_bytes))
        except:
            # Pokud je to CSV/Text, detekujeme kódování
            res = chardet.detect(file_bytes)
            enc = res['encoding'] if res['encoding'] else 'cp1250'
            text_data = file_bytes.decode(enc, errors='replace')
            
            # Přečteme data s automatickou detekcí oddělovače (tabulátor/středník/čárka)
            df = pd.read_csv(io.StringIO(text_data), sep=None, engine='python', skipinitialspace=True)

        # Vyčištění názvů sloupců (odstranění mezer a uvozovek)
        df.columns = [str(c).strip().replace('"', '') for c in df.columns]

        # 2. KONTROLA SLOUPCE
        if 'Vystaveno' not in df.columns:
            st.error(f"Sloupec 'Vystaveno' nebyl nalezen. Sloupce v souboru jsou: {', '.join(df.columns)}")
            st.stop()

        # 3. FILTRACE ČASU (10:00 - 12:00 druhý den)
        df['dt'] = pd.to_datetime(df['Vystaveno'], dayfirst=True, errors='coerce')
        df = df.dropna(subset=['dt'])
        
        min_date = df['dt'].min().date()
        start_t = datetime.combine(min_date, time(10, 0, 0))
        end_t = datetime.combine(min_date + pd.Timedelta(days=1), time(12, 0, 0))
        
        df_f = df[(df['dt'] >= start_t) & (df['dt'] <= end_t)].copy()

        # 4. VÝBĚR A PŘEJMENOVÁNÍ SLOUPCŮ (podle tvého raw souboru)
        # Používáme přesné názvy, které jsi poslal
        mapping = {
            "Vystaveno": "Vystaveno",
            "Stav": "Stav",
            "Číslo": "Číslo",
            "Variabilní symbol": "Variabilní symbol",
            "Forma úhrady": "Forma úhrady",
            "DUZP": "Splatnost",
            "Základ 0%": "Základ 0%",
            "Základ - snížená sazba 12% (15%)": "Základ 12%",
            "DPH - snížená sazba 12% (15%)": "DPH 12%",
            "Základ - základní sazba 21%": "Základ 21%",
            "DPH - základní sazba 21%": "DPH 21%",
            "Celkem bez DPH": "Celkem bez DPH",
            "Celkem s DPH": "Celkem s DPH"
        }
        
        available = [c for c in mapping.keys() if c in df_f.columns]
        df_final = df_f[available].rename(columns=mapping)

        # Převod na čísla
        for col in df_final.columns:
            if any(x in col for x in ["Základ", "DPH", "Celkem"]):
                df_final[col] = pd.to_numeric(df_final[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)

        # 5. EXPORT DO EXCELU
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb, ws = writer.book, writer.book.add_worksheet('Report')
            f_bold, f_num = wb.add_format({'bold': True}), wb.add_format({'num_format': '#,##0.00'})
            
            d1, d2 = start_t.strftime("%d.%m."), end_t.strftime("%d.%m.%Y")
            ws.write('B1', f"Tržba ze dne {d1} - {d2}", f_bold)
            
            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)
            
            n_rows = len(df_final)
            if n_rows > 0:
                fr, lr, sr = 4, 4 + n_rows - 1, 4 + n_rows + 1
                # SUMIF - Forma úhrady (sloupec E), Celkem s DPH (sloupec M/podle počtu sloupců)
                # Pro jistotu použijeme pevné vzorce na základě tvé struktury
                ws.write(sr, 2, "Hotovost celkem:", f_bold)
                ws.write_formula(sr, 3, f'=SUMIF(E{fr}:E{lr}, "*Hotově*", M{fr}:M{lr})', f_num)
                ws.write(sr + 1, 2, "Karty celkem:", f_bold)
                ws.write_formula(sr + 1, 3, f'=SUMIF(E{fr}:E{lr}, "*Kartou*", M{fr}:M{lr})', f_num)

        st.success("✅ Report vygenerován!")
        st.download_button(label="📥 Stáhnout Excel", data=output.getvalue(), file_name=f"Report_{d1}{d2}.xlsx")
    except Exception as e:
        st.error(f"Chyba: {e}")

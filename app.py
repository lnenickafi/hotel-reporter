import streamlit as st
import pandas as pd
import io
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")
st.title("🏨 Hotelový Reportér")
st.write("Nahrajte raw export (XLS)")

uploaded_file = st.file_uploader("Soubor XLS", type=["xls"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()
        
        # 1. NAČTENÍ BINÁRNÍHO XLS
        # Pro staré XLS je nejlepší xlrd (v pandas engine='xlrd')
        # Načteme vše a pak budeme hledat hlavičku
        df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='xlrd')

        # 2. HLEDÁNÍ HLAVIČKY (Vystaveno)
        header_idx = None
        for i in range(len(df_raw)):
            row = [str(x).strip() for x in df_raw.iloc[i].values]
            if "Vystaveno" in row:
                header_idx = i
                df_raw.columns = row
                break
        
        if header_idx is None:
            st.error("Sloupec 'Vystaveno' nebyl nalezen. Zkontrolujte, zda jde o správný export.")
            st.stop()

        # Oříznutí dat pod hlavičkou
        df = df_raw.iloc[header_idx + 1:].reset_index(drop=True)
        
        # 3. FILTRACE ČASU (10:00 - 12:00)
        # Převod sloupce Vystaveno na datum (Excel ho často už jako datum má)
        df['dt_obj'] = pd.to_datetime(df['Vystaveno'], errors='coerce')
        df = df.dropna(subset=['dt_obj'])
        
        if df.empty:
            st.error("V souboru nebyla nalezena žádná platná data.")
            st.stop()

        min_date = df['dt_obj'].min().date()
        start_t = datetime.combine(min_date, time(10, 0, 0))
        end_t = datetime.combine(min_date + pd.Timedelta(days=1), time(12, 0, 0))
        
        df_f = df[(df['dt_obj'] >= start_t) & (df['dt_obj'] <= end_t)].copy()

        # 4. VÝBĚR SLOUPCŮ (přesně podle tvého binárního výpisu)
        # Názvy v mappingu musí odpovídat tomu, co jsi poslal
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
        
        # Vybereme jen ty, které v souboru skutečně jsou
        available_cols = [c for c in mapping.keys() if c in df_f.columns]
        df_final = df_f[available_cols].rename(columns=mapping)

        # Převod čísel
        for col in df_final.columns:
            if any(x in col for x in ["Základ", "DPH", "Celkem"]):
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

        # 5. GENEROVÁNÍ NOVÉHO EXCELU
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb, ws = writer.book, writer.book.add_worksheet('Report')
            f_bold, f_num = wb.add_format({'bold': True}), wb.add_format({'num_format': '#,##0.00'})
            
            d1_str = start_t.strftime("%d.%m.")
            d2_str = end_t.strftime("%d.%m.%Y")
            ws.write('B1', f"Report tržeb {d1_str} (10:00) - {d2_str} (12:00)", f_bold)
            
            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)
            
            # Automatické součty pro Formu úhrady (sloupec E) a Celkem s DPH (sloupec M)
            n_rows = len(df_final)
            if n_rows > 0:
                fr, lr, sr = 4, 4 + n_rows - 1, 4 + n_rows + 1
                ws.write(sr, 2, "Hotovost celkem:", f_bold)
                ws.write_formula(sr, 3, f'=SUMIF(E{fr}:E{lr}, "*Hotově*", M{fr}:M{lr})', f_num)
                ws.write(sr + 1, 2, "Karty celkem:", f_bold)
                ws.write_formula(sr + 1, 3, f'=SUMIF(E{fr}:E{lr}, "*Kartou*", M{fr}:M{lr})', f_num)

        st.success("✅ Report úspěšně vytvořen!")
        st.download_button(
            label="📥 Stáhnout upravený Excel",
            data=output.getvalue(),
            file_name=f"Report_{d1_str}{d2_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Chyba při zpracování: {e}")

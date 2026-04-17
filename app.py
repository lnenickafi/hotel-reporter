import streamlit as st
import pandas as pd
import io
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")

st.title("🏨 Hotelový Reportér")

with st.expander("💡 INSTRUKCE: Jak připravit soubor?", expanded=True):
    st.markdown("""
    1. Otevřete export v **Excelu**.
    2. Klikněte na **Soubor** -> **Uložit jako**.
    3. Vyberte formát **Sešit Excel (*.xlsx)**.
    4. Tento nový soubor nahrajte sem.
    """)

uploaded_file = st.file_uploader("Nahrajte převedený soubor (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # 1. NAČTENÍ XLSX
        df_raw = pd.read_excel(uploaded_file, header=None)

        # 2. HLEDÁNÍ HLAVIČKY
        header_idx = None
        for i in range(len(df_raw)):
            row = [str(x).strip() for x in df_raw.iloc[i].values]
            if "Vystaveno" in row:
                header_idx = i
                # Nastavíme vyčištěné názvy sloupců
                df_raw.columns = row
                break
        
        if header_idx is None:
            st.error("❌ V souboru nebyl nalezen sloupec 'Vystaveno'.")
            st.stop()

        # Oříznutí a vyčištění
        df = df_raw.iloc[header_idx + 1:].reset_index(drop=True)
        df.columns = [str(c).strip() for c in df.columns]

        # 3. OPRAVA CHYBY: Převod data (Pojistka proti 'arg must be a list...')
        # Najdeme první sloupec, který se jmenuje přesně 'Vystaveno'
        if 'Vystaveno' in df.columns:
            # Zajistíme, že bereme jen JEDEN sloupec (i kdyby jich bylo víc)
            vystaveno_data = df['Vystaveno']
            if isinstance(vystaveno_data, pd.DataFrame):
                vystaveno_data = vystaveno_data.iloc[:, 0]
            
            df['dt_obj'] = pd.to_datetime(vystaveno_data, dayfirst=True, errors='coerce')
        else:
            st.error("Sloupec 'Vystaveno' se nepodařilo po načtení identifikovat.")
            st.stop()

        df = df.dropna(subset=['dt_obj'])
        
        if df.empty:
            st.warning("⚠️ Žádná platná data k seřazení.")
            st.stop()

        # Filtrace 10:00 - 12:00
        min_date = df['dt_obj'].min().date()
        start_t = datetime.combine(min_date, time(10, 0, 0))
        end_t = datetime.combine(min_date + pd.Timedelta(days=1), time(12, 0, 0))
        df_f = df[(df['dt_obj'] >= start_t) & (df['dt_obj'] <= end_t)].copy()

        # 4. MAPOVÁNÍ SLOUPCŮ
        mapping_rules = {
            "Vystaveno": "Vystaveno", "Stav": "Stav", "Číslo": "Číslo",
            "Variabilní symbol": "Variabilní symbol", "Forma úhrady": "Forma úhrady",
            "Splatnost": "Splatnost", "Základ 0%": "Základ 0%",
            "12%": "DPH 12%", "21%": "DPH 21%",
            "Celkem bez DPH": "Celkem bez DPH", "Celkem s DPH": "Celkem s DPH"
        }
        
        final_mapping = {}
        for real_col in df_f.columns:
            for key, target_name in mapping_rules.items():
                if key.lower() in str(real_col).lower():
                    final_mapping[real_col] = target_name
                    break
        
        # Výběr sloupců, které existují
        cols_to_keep = [c for c in df_f.columns if c in final_mapping]
        df_final = df_f[cols_to_keep].rename(columns=final_mapping)

        # Čištění čísel
        for col in df_final.columns:
            if any(x in col for x in ["Základ", "DPH", "Celkem"]):
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

        # 5. EXPORT
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb, ws = writer.book, writer.book.add_worksheet('Report')
            f_bold, f_num = wb.add_format({'bold': True}), wb.add_format({'num_format': '#,##0.00'})
            
            d1, d2 = start_t.strftime("%d.%m."), end_t.strftime("%d.%m.%Y")
            ws.write('B1', f"Tržba {d1} - {d2}", f_bold)
            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)
            
            n_rows = len(df_final)
            if n_rows > 0:
                fr, lr, sr = 4, 4 + n_rows - 1, 4 + n_rows + 1
                try:
                    cols = list(df_final.columns)
                    c_forma = chr(65 + cols.index("Forma úhrady"))
                    c_celkem = chr(65 + cols.index("Celkem s DPH"))
                    ws.write(sr, 2, "Hotovost celkem:", f_bold)
                    ws.write_formula(sr, 3, f'=SUMIF({c_forma}{fr}:{c_forma}{lr}, "*Hotově*", {c_celkem}{fr}:{c_celkem}{lr})', f_num)
                    ws.write(sr + 1, 2, "Karty celkem:", f_bold)
                    ws.write_formula(sr + 1, 3, f'=SUMIF({c_forma}{fr}:{c_forma}{lr}, "*Kartou*", {c_celkem}{fr}:{c_celkem}{lr})', f_num)
                except: pass

        st.success("✅ Report vygenerován!")
        st.download_button(label="📥 Stáhnout", data=output.getvalue(), file_name=f"Report_{d1}{d2}.xlsx")

    except Exception as e:
        st.error(f"Chyba: {e}")

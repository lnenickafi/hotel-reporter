import streamlit as st
import pandas as pd
import io
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")

# Nadpis a Instrukce
st.title("🏨 Hotelový Reportér")

with st.expander("💡 INSTRUKCE: Jak připravit soubor, aby fungoval?", expanded=True):
    st.markdown("""
    Protože hotelový systém generuje starý a chybný typ souboru, je nutné udělat tento krátký krok:
    1. Otevřete stažený export v **Excelu**.
    2. Klikněte na **Soubor** -> **Uložit jako**.
    3. Vyberte formát **Sešit Excel (*.xlsx)**.
    4. Tento nový soubor nahrajte sem do aplikace.
    """)

uploaded_file = st.file_uploader("Nahrajte převedený soubor (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # 1. NAČTENÍ XLSX
        # Načteme bez hlavičky, abychom ji našli (kdyby tam byl ten řádek "Tržba ze dne...")
        df_raw = pd.read_excel(uploaded_file, header=None)

        # 2. HLEDÁNÍ HLAVIČKY (Sloupec Vystaveno)
        header_idx = None
        for i in range(len(df_raw)):
            row = [str(x).strip() for x in df_raw.iloc[i].values]
            if "Vystaveno" in row:
                header_idx = i
                df_raw.columns = row
                break
        
        if header_idx is None:
            st.error("❌ V souboru nebyl nalezen sloupec 'Vystaveno'. Ujistěte se, že nahráváte správný export.")
            st.stop()

        # Oříznutí dat pod hlavičkou
        df = df_raw.iloc[header_idx + 1:].reset_index(drop=True)
        
        # Vyčištění názvů sloupců od mezer
        df.columns = [str(c).strip() for c in df.columns]

        # 3. FILTRACE ČASU (10:00 - 12:00 druhý den)
        # Excel už v XLSX formátu obvykle drží datumy jako objekty
        df['dt_obj'] = pd.to_datetime(df['Vystaveno'], dayfirst=True, errors='coerce')
        df = df.dropna(subset=['dt_obj'])
        
        if df.empty:
            st.warning("⚠️ V souboru nebyla nalezena žádná platná data k seřazení.")
            st.stop()

        # Určení rozmezí
        min_date = df['dt_obj'].min().date()
        start_t = datetime.combine(min_date, time(10, 0, 0))
        end_t = datetime.combine(min_date + pd.Timedelta(days=1), time(12, 0, 0))
        
        df_f = df[(df['dt_obj'] >= start_t) & (df['dt_obj'] <= end_t)].copy()

        # 4. MAPOVÁNÍ SLOUPCŮ (hledání klíčových slov v názvech)
        mapping_rules = {
            "Vystaveno": "Vystaveno",
            "Stav": "Stav",
            "Číslo": "Číslo",
            "Variabilní symbol": "Variabilní symbol",
            "Forma úhrady": "Forma úhrady",
            "DUZP": "Splatnost",
            "Základ 0%": "Základ 0%",
            "12%": "DPH 12%",
            "21%": "DPH 21%",
            "Celkem bez DPH": "Celkem bez DPH",
            "Celkem s DPH": "Celkem s DPH"
        }
        
        final_mapping = {}
        for real_col in df_f.columns:
            for key, target_name in mapping_rules.items():
                if key.lower() in real_col.lower():
                    final_mapping[real_col] = target_name
                    break
        
        df_final = df_f[list(final_mapping.keys())].rename(columns=final_mapping)

        # Převod čísel (očištění od případných textů/měn)
        for col in df_final.columns:
            if any(x in col for x in ["Základ", "DPH", "Celkem"]):
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

        # 5. EXPORT DO NOVÉHO EXCELU
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb = writer.book
            ws = wb.add_worksheet('Report')
            f_bold = wb.add_format({'bold': True})
            f_num = wb.add_format({'num_format': '#,##0.00'})
            
            d1 = start_t.strftime("%d.%m.")
            d2 = end_t.strftime("%d.%m.%Y")
            ws.write('B1', f"Tržba {d1} (10:00) až {d2} (12:00)", f_bold)
            
            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)
            
            # Sumy na konci
            n_rows = len(df_final)
            if n_rows > 0:
                fr, lr = 4, 4 + n_rows - 1
                sr = lr + 2
                
                # Zkusíme najít písmena sloupců dynamicky
                cols = list(df_final.columns)
                try:
                    c_forma = chr(65 + cols.index("Forma úhrady"))
                    c_celkem = chr(65 + cols.index("Celkem s DPH"))
                    
                    ws.write(sr, 2, "Hotovost celkem:", f_bold)
                    ws.write_formula(sr, 3, f'=SUMIF({c_forma}{fr}:{c_forma}{lr}, "*Hotově*", {c_celkem}{fr}:{c_celkem}{lr})', f_num)
                    ws.write(sr + 1, 2, "Karty celkem:", f_bold)
                    ws.write_formula(sr + 1, 3, f'=SUMIF({c_forma}{fr}:{c_forma}{lr}, "*Kartou*", {c_celkem}{fr}:{c_celkem}{lr})', f_num)
                except:
                    pass

        st.success("✅ Report úspěšně vygenerován z XLSX!")
        st.download_button(
            label="📥 Stáhnout upravenou uzávěrku",
            data=output.getvalue(),
            file_name=f"Report_{d1}{d2}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Chyba při zpracování: {e}")

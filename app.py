import streamlit as st
import pandas as pd
import io
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")
st.title("🏨 Hotelový Reportér")

st.markdown("""
### 🚀 Jak nahrát export?
1. **Ideální cesta:** V systému zvolte export **CSV (oddělovač čárka)** a **vypněte souhrny/filtry**. Nahrajte přímo stažený soubor.
2. **Záložní cesta:** Pokud máte `.xls`, uložte jej v Excelu jako **Sešit Excel (.xlsx)** a nahrajte ten.
""")

uploaded_file = st.file_uploader("Nahrajte soubor (CSV nebo XLSX)", type=["csv", "xlsx", "xls"])

if uploaded_file:
    try:
        # 1. NAČTENÍ (univerzální pro CSV i Excel)
        if uploaded_file.name.lower().endswith('.csv'):
            # Zkusíme načíst CSV (čárka i středník)
            try:
                df_raw = pd.read_csv(uploaded_file, sep=',', header=None, encoding='utf-8')
            except:
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, sep=';', header=None, encoding='cp1250')
        else:
            df_raw = pd.read_excel(uploaded_file, header=None)

        # 2. DETEKCE SLOUPCŮ PODLE POZICE (nejstabilnější metoda)
        header_row_idx = None
        col_idx = {}
        
        search_map = {
            "vyst": "vystaveno", "stav": "stav", "cislo": "číslo",
            "var": "variabil", "forma": "forma úhrady", "duzp": "splatnost",
            "z0": "základ 0%", "z12": "12%", "z21": "21%",
            "bez": "bez dph", "s_dph": "s dph"
        }

        for i in range(len(df_raw)):
            row = [str(x).lower().strip() for x in df_raw.iloc[i].values]
            if "vystaveno" in row:
                header_row_idx = i
                for key, term in search_map.items():
                    for idx, cell in enumerate(row):
                        if term in cell:
                            col_idx[key] = idx
                break

        if header_row_idx is None:
            st.error("❌ V souboru nebyl nalezen sloupec 'Vystaveno'.")
            st.stop()

        # 3. ZPRACOVÁNÍ DAT
        df_data = df_raw.iloc[header_row_idx + 1:].copy().reset_index(drop=True)
        
        # Převod data
        v_pos = col_idx.get("vyst")
        df_data['dt_fixed'] = pd.to_datetime(df_data.iloc[:, v_pos], dayfirst=True, errors='coerce')
        df_data = df_data.dropna(subset=['dt_fixed'])

        # Filtrace času (10:00 - 12:00)
        min_d = df_data['dt_fixed'].min().date()
        st_range = datetime.combine(min_d, time(10, 0, 0))
        en_range = datetime.combine(min_d + pd.Timedelta(days=1), time(12, 0, 0))
        df_f = df_data[(df_data['dt_fixed'] >= st_range) & (df_data['dt_fixed'] <= en_range)].copy()

        # 4. TVORBA TABULKY A OPRAVA 'PR'
        final_list = []
        for _, row in df_f.iterrows():
            # Oprava čísla (vždy PR + odstranění .0)
            raw_c = str(row.iloc[col_idx["cislo"]]).split('.')[0].strip()
            clean_c = raw_c if raw_c.upper().startswith("PR") else f"PR{raw_c}"
            
            # Vyčištění Variabilního symbolu
            var_sym = str(row.iloc[col_idx["var"]]).split('.')[0].strip()

            final_list.append({
                "Vystaveno": row.iloc[col_idx["vyst"]],
                "Stav": row.iloc[col_idx["stav"]],
                "Číslo": clean_c,
                "Variabilní symbol": var_sym,
                "Forma úhrady": row.iloc[col_idx["forma"]],
                "Splatnost": row.iloc[col_idx["duzp"]],
                "Základ 0%": pd.to_numeric(str(row.iloc[col_idx["z0"]]).replace(',', '.'), errors='coerce'),
                "DPH - 12%": pd.to_numeric(str(row.iloc[col_idx["z12"]]).replace(',', '.'), errors='coerce'),
                "DPH 21%": pd.to_numeric(str(row.iloc[col_idx["z21"]]).replace(',', '.'), errors='coerce'),
                "Celkem bez DPH": pd.to_numeric(str(row.iloc[col_idx["bez"]]).replace(',', '.'), errors='coerce'),
                "Celkem s DPH": pd.to_numeric(str(row.iloc[col_idx["s_dph"]]).replace(',', '.'), errors='coerce')
            })

        df_final = pd.DataFrame(final_list).fillna(0)

        # 5. EXPORT
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb, ws = writer.book, writer.book.add_worksheet('Report')
            f_bold, f_num = wb.add_format({'bold': True}), wb.add_format({'num_format': '#,##0.00'})
            
            d1, d2 = st_range.strftime("%d.%m."), en_range.strftime("%d.%m. %Y")
            ws.write('B1', f"Tržba ze dne {d1} - {d2}", f_bold)
            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)

            # SOUČTY (SUMIF)
            n = len(df_final)
            if n > 0:
                fr, lr, sr = 4, 4 + n - 1, 4 + n + 1
                # E=Forma úhrady, K=Celkem s DPH
                ws.write(sr, 2, "Hotovost celkem:", f_bold)
                ws.write_formula(sr, 3, f'=SUMIF(E{fr}:E{lr}, "*Hotově*", K{fr}:K{lr})', f_num)
                ws.write(sr + 1, 2, "Kreditní karty celkem:", f_bold)
                ws.write_formula(sr + 1, 3, f'=SUMIF(E{fr}:E{lr}, "*Kartou*", K{fr}:K{lr})', f_num)
            
            ws.set_column('A:K', 18)

        st.success("✅ Report úspěšně vygenerován!")
        st.download_button("📥 Stáhnout hotový report", output.getvalue(), f"Uzaverka_{d1}_{d2}.xlsx")

    except Exception as e:
        st.error(f"Chyba: {e}")

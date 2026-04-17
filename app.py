import streamlit as st
import pandas as pd
import io
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")

# --- HLAVIČKA A INSTRUKCE ---
st.title("🏨 Hotelový Reportér")

st.markdown("""
### 💡 Jak nahrát data správně?
Pro správné fungování reportu postupujte takto:
1. Vyexportujte data z hotelového systému (soubor `.xls`).
2. Otevřete tento soubor v **Excelu**.
3. Klikněte na **Soubor** -> **Uložit jako**.
4. Zvolte formát **Sešit Excel (*.xlsx)** a uložte.
5. Tento nový soubor nahrajte níže.
""")

uploaded_file = st.file_uploader("Nahrajte převedený .xlsx soubor", type=["xlsx"])

if uploaded_file:
    try:
        # 1. NAČTENÍ (bez hlaviček pro totální kontrolu)
        df_raw = pd.read_excel(uploaded_file, header=None)

        # 2. DETEKCE POZIC SLOUPCŮ
        header_row_idx = None
        col_map = {}
        
        # Klíčová slova pro identifikaci sloupců
        search_terms = {
            "vyst": "vystaveno",
            "stav": "stav",
            "cislo": "číslo",
            "var": "variabil",
            "forma": "forma úhrady",
            "duzp": "splatnost",
            "z0": "základ 0%",
            "z12": "12%",
            "z21": "21%",
            "bez": "bez dph",
            "s_dph": "s dph"
        }

        for i in range(len(df_raw)):
            row = [str(x).lower().strip() for x in df_raw.iloc[i].values]
            if "vystaveno" in row:
                header_row_idx = i
                for key, term in search_terms.items():
                    for idx, cell in enumerate(row):
                        if term in cell:
                            col_map[key] = idx
                break

        if header_row_idx is None:
            st.error("❌ V souboru nebyla nalezena tabulka (sloupec Vystaveno).")
            st.stop()

        # 3. ZPRACOVÁNÍ DAT
        df_data = df_raw.iloc[header_row_idx + 1:].copy().reset_index(drop=True)
        
        # Převod data a času
        v_idx = col_map.get("vyst")
        df_data['dt_fixed'] = pd.to_datetime(df_data.iloc[:, v_idx], dayfirst=True, errors='coerce')
        df_data = df_data.dropna(subset=['dt_fixed'])

        # Filtrace času (10:00 - 12:00 druhý den)
        min_d = df_data['dt_fixed'].min().date()
        st_range = datetime.combine(min_d, time(10, 0, 0))
        en_range = datetime.combine(min_d + pd.Timedelta(days=1), time(12, 0, 0))
        df_f = df_data[(df_data['dt_fixed'] >= st_range) & (df_data['dt_fixed'] <= en_range)].copy()

        # 4. TVORBA FINÁLNÍ TABULKY S OPRAVAMI
        final_rows = []
        for _, row in df_f.iterrows():
            # Oprava čísla dokladu (přidání PR pokud chybí)
            raw_val = str(row.iloc[col_map["cislo"]])
            clean_cislo = raw_val if raw_val.startswith("PR") else f"PR{raw_val}"
            # Odstranění .0 na konci pokud se to načetlo jako číslo
            clean_cislo = clean_cislo.replace(".0", "")

            new_row = {
                "Vystaveno": row.iloc[col_map["vyst"]],
                "Stav": row.iloc[col_map["stav"]],
                "Číslo": clean_cislo,
                "Variabilní symbol": str(row.iloc[col_map["var"]]).replace(".0", ""),
                "Forma úhrady": row.iloc[col_map["forma"]],
                "Splatnost": row.iloc[col_map["duzp"]],
                "Základ 0%": pd.to_numeric(row.iloc[col_map["z0"]], errors='coerce'),
                "DPH - 12%": pd.to_numeric(row.iloc[col_map["z12"]], errors='coerce'),
                "DPH 21%": pd.to_numeric(row.iloc[col_map["z21"]], errors='coerce'),
                "Celkem bez DPH": pd.to_numeric(row.iloc[col_map["bez"]], errors='coerce'),
                "Celkem s DPH": pd.to_numeric(row.iloc[col_map["s_dph"]], errors='coerce')
            }
            final_rows.append(new_row)

        df_final = pd.DataFrame(final_rows).fillna(0)

        # 5. EXPORT DO EXCELU
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb, ws = writer.book, writer.book.add_worksheet('Report')
            f_bold = wb.add_format({'bold': True, 'font_size': 12})
            f_num = wb.add_format({'num_format': '#,##0.00'})
            
            d1, d2 = st_range.strftime("%d.%m."), en_range.strftime("%d.%m. %Y")
            ws.write('B1', f"Tržba ze dne {d1} - {d2}", f_bold)

            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)

            # SOUČTY (SUMIF)
            n = len(df_final)
            if n > 0:
                fr, lr, sr = 4, 4 + n - 1, 4 + n + 1
                # E = Forma úhrady, K = Celkem s DPH
                ws.write(sr, 2, "Hotovost celkem:", f_bold)
                ws.write_formula(sr, 3, f'=SUMIF(E{fr}:E{lr}, "*Hotově*", K{fr}:K{lr})', f_num)
                ws.write(sr + 1, 2, "Kreditní karty celkem:", f_bold)
                ws.write_formula(sr + 1, 3, f'=SUMIF(E{fr}:E{lr}, "*Kartou*", K{fr}:K{lr})', f_num)
            
            ws.set_column('A:K', 18)

        st.success(f"✅ Report vygenerován!")
        st.download_button("📥 Stáhnout hotový report", output.getvalue(), f"Uzaverka_{d1}_{d2}.xlsx")

    except Exception as e:
        st.error(f"Chyba: {e}")

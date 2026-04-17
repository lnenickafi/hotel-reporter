import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime, time

st.set_page_config(page_title="Reportér úzávěrky", page_icon="🏨")

st.title("🏨 Uzávěrka tržeb")

st.markdown("""
### Kdyby cokoliv nefungovalo, napsat Filipovi!
💡 Jak nahrát data správně?
Pro správné fungování reportu postupujte takto:
1. Vyexportujte data z hotelového systému (soubor `.xls`).
2. Otevřete tento soubor v **Excelu**.
3. Klikněte na **Soubor** -> **Uložit jako**.
4. Zvolte formát **Sešit Excel (*.xlsx)** a uložte.
5. Tento nový soubor nahrajte níže.

- je nutné takto postupovat a převést to do XLSX, jinak to nefunguje. -
- Časy dokladů to ořezává od 10:00 počátečního dne do 12:00 dalšího dne, kdyby bylo potřeba upravit, dát mi vědět -
""")

uploaded_file = st.file_uploader("Nahrajte .xlsx soubor", type=["xlsx"])

if uploaded_file:
    try:
        # 1. NAČTENÍ (Načteme vše jako text, abychom předešli chybám formátu)
        df_raw = pd.read_excel(uploaded_file, header=None, dtype=str)

        # 2. IDENTIFIKACE SLOUPCŮ (hledáme řádek s hlavičkou)
        header_row_idx = None
        col_idx = {}
        search_terms = {
            "vyst": "vystaveno", "stav": "stav", "cislo": "číslo",
            "var": "variabil", "forma": "forma úhrady", "duzp": "splatnost",
            "z0": "základ 0%", "z12": "12%", "z21": "21%",
            "bez": "bez dph", "s_dph": "s dph"
        }

        for i in range(len(df_raw)):
            row = [str(x).lower().strip() for x in df_raw.iloc[i].values]
            if "vystaveno" in row:
                header_row_idx = i
                for key, term in search_terms.items():
                    for idx, cell in enumerate(row):
                        if term in cell:
                            col_idx[key] = idx
                break

        if header_row_idx is None:
            st.error("❌ V souboru nebyla nalezena tabulka (sloupec Vystaveno).")
            st.stop()

        # 3. FILTRACE A ČIŠTĚNÍ
        df_data = df_raw.iloc[header_row_idx + 1:].copy().reset_index(drop=True)
        
        # Převedeme datum (ošetříme i textové formáty)
        v_pos = col_idx.get("vyst")
        df_data['dt_fixed'] = pd.to_datetime(df_data.iloc[:, v_pos], dayfirst=True, errors='coerce')
        
        # Klíčový krok: Odstraníme řádky, které nemají datum (to jsou ty souhrny a filtry na konci!)
        df_data = df_data.dropna(subset=['dt_fixed'])

        if df_data.empty:
            st.warning("⚠️ Žádná platná data k seřazení.")
            st.stop()

        # Filtrace 10:00 (den 1) - 12:00 (den 2)
        min_d = df_data['dt_fixed'].min().date()
        st_range = datetime.combine(min_d, time(10, 0, 0))
        en_range = datetime.combine(min_d + pd.Timedelta(days=1), time(12, 0, 0))
        df_f = df_data[(df_data['dt_fixed'] >= st_range) & (df_data['dt_fixed'] <= en_range)].copy()

        # 4. AGRESIVNÍ ČIŠTĚNÍ ČÍSEL
        def parse_cz_number(val):
            if pd.isna(val) or val == "" or val == "nan": return 0.0
            # Odstraníme uvozovky, mezery, měny (CZK)
            s = str(val).replace('"', '').replace(' ', '').replace('\xa0', '')
            s = s.replace('CZK', '').replace('Kč', '').strip()
            # Nahradíme čárku tečkou
            s = s.replace(',', '.')
            # Vytáhneme jen číselné znaky (včetně mínus a tečky)
            s = re.sub(r'[^0-9.-]', '', s)
            try:
                return float(s)
            except:
                return 0.0

        # 5. TVORBA FINÁLNÍ TABULKY
        final_list = []
        for _, row in df_f.iterrows():
            # Číslo dokladu s PR
            raw_c = str(row.iloc[col_idx["cislo"]]).replace(".0", "").replace('"', '').strip()
            clean_c = raw_c if raw_c.upper().startswith("PR") else f"PR{raw_c}"
            
            # Variabilní symbol (bez .0)
            var_s = str(row.iloc[col_idx["var"]]).replace(".0", "").replace('"', '').strip()

            final_list.append({
                "Vystaveno": row.iloc[col_idx["vyst"]],
                "Stav": row.iloc[col_idx["stav"]],
                "Číslo": clean_c,
                "Variabilní symbol": var_s,
                "Forma úhrady": row.iloc[col_idx["forma"]],
                "Splatnost": row.iloc[col_idx["duzp"]],
                "Základ 0%": parse_cz_number(row.iloc[col_idx["z0"]]),
                "DPH - 12%": parse_cz_number(row.iloc[col_idx["z12"]]),
                "DPH 21%": parse_cz_number(row.iloc[col_idx["z21"]]),
                "Celkem bez DPH": parse_cz_number(row.iloc[col_idx["bez"]]),
                "Celkem s DPH": parse_cz_number(row.iloc[col_idx["s_dph"]])
            })

        df_final = pd.DataFrame(final_list)

        # 6. EXPORT DO EXCELU
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb, ws = writer.book, writer.book.add_worksheet('Report')
            f_bold, f_num = wb.add_format({'bold': True}), wb.add_format({'num_format': '#,##0.00'})
            
            d1_s, d2_s = st_range.strftime("%d.%m."), en_range.strftime("%d.%m. %Y")
            ws.write('B1', f"Tržba ze dne {d1_s} - {d2_s}", f_bold)
            
            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)

            # SOUČTY (E = Forma úhrady, K = Celkem s DPH)
            n = len(df_final)
            if n > 0:
                fr, lr, sr = 4, 4 + n - 1, 4 + n + 1
                ws.write(sr, 2, "Hotově:", f_bold)
                ws.write_formula(sr, 3, f'=SUMIF(E{fr}:E{lr}, "*Hotově*", K{fr}:K{lr})', f_num)
                ws.write(sr + 1, 2, "Kred. kartou:", f_bold)
                ws.write_formula(sr + 1, 3, f'=SUMIF(E{fr}:E{lr}, "*Kartou*", K{fr}:K{lr})', f_num)
            
            ws.set_column('A:K', 18)

        st.success(f"✅ Uzávěrka pro {d1_s} - {d2_s} vygenerována!")
        st.download_button("📥 Stáhnout hotový report", output.getvalue(), f"Report_{d1_s}{d2_s}.xlsx")

    except Exception as e:
        st.error(f"Chyba: {e}")

import streamlit as st
import pandas as pd
import io
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")
st.title("🏨 Hotelový Reportér")

st.info("Nahrajte převedený soubor .xlsx (uložený v Excelu jako Sešit Excel).")

uploaded_file = st.file_uploader("Nahrajte .xlsx soubor", type=["xlsx"])

if uploaded_file:
    try:
        # 1. NAČTENÍ
        df_raw = pd.read_excel(uploaded_file, header=None)

        # 2. DETEKCE STRUKTURY (podle pozic)
        header_row_idx = None
        col_indices = {}
        
        # Klíčová slova pro mapování (hledáme v raw datech)
        targets = {
            "vyst": "Vystaveno",
            "stav": "Stav",
            "cislo": "Číslo",
            "var": "Variabilní symbol",
            "forma": "Forma úhrady",
            "duzp": "Splatnost",
            "z0": "Základ 0%",
            "z12": "DPH - 12%",
            "z21": "DPH 21%",
            "bez": "Celkem bez DPH",
            "s_dph": "Celkem s DPH"
        }

        for i in range(len(df_raw)):
            row_list = [str(x).lower().strip() for x in df_raw.iloc[i].values]
            if any("vystaveno" in str(x) for x in row_list):
                header_row_idx = i
                # Najdeme indexy pro každý sloupec
                for col_idx, cell_val in enumerate(row_list):
                    if "vystaveno" in cell_val: col_indices["vyst"] = col_idx
                    if "stav" in cell_val: col_indices["stav"] = col_idx
                    if "číslo" in cell_val: col_indices["cislo"] = col_idx
                    if "variabil" in cell_val or "var." in cell_val: col_indices["var"] = col_idx
                    if "forma" in cell_val: col_indices["forma"] = col_idx
                    if "splat" in cell_val or "duzp" in cell_val: col_indices["duzp"] = col_idx
                    if "0%" in cell_val: col_indices["z0"] = col_idx
                    if "12%" in cell_val: col_indices["z12"] = col_idx
                    if "21%" in cell_val: col_indices["z21"] = col_idx
                    if "bez dph" in cell_val: col_indices["bez"] = col_idx
                    if "s dph" in cell_val: col_indices["s_dph"] = col_idx
                break

        if header_row_idx is None or "vyst" not in col_indices:
            st.error("❌ Nepodařilo se najít sloupec 'Vystaveno'.")
            st.stop()

        # 3. ZPRACOVÁNÍ DAT
        df_data = df_raw.iloc[header_row_idx + 1:].copy().reset_index(drop=True)
        
        # Převod data
        vyst_idx = col_indices["vyst"]
        df_data['dt_fixed'] = pd.to_datetime(df_data.iloc[:, vyst_idx], dayfirst=True, errors='coerce')
        df_data = df_data.dropna(subset=['dt_fixed'])

        # Filtrace 10:00 - 12:00
        min_d = df_data['dt_fixed'].min().date()
        st_range = datetime.combine(min_d, time(10, 0, 0))
        en_range = datetime.combine(min_d + pd.Timedelta(days=1), time(12, 0, 0))
        df_f = df_data[(df_data['dt_fixed'] >= st_range) & (df_data['dt_fixed'] <= en_range)].copy()

        # SESTAVENÍ TABULKY V PŘESNÉM POŘADÍ
        # (Definujeme sloupce, které chceme mít ve finálním Excelu)
        final_order = [
            ("vyst", "Vystaveno"), ("stav", "Stav"), ("cislo", "Číslo"), 
            ("var", "Variabilní symbol"), ("forma", "Forma úhrady"), ("duzp", "Splatnost"),
            ("z0", "Základ 0%"), ("z12", "DPH - 12%"), ("z21", "DPH 21%"),
            ("bez", "Celkem bez DPH"), ("s_dph", "Celkem s DPH")
        ]

        export_dict = {}
        for key, name in final_order:
            if key in col_indices:
                export_dict[name] = df_f.iloc[:, col_indices[key]]

        df_final = pd.DataFrame(export_dict)

        # Čištění čísel
        num_cols = ["Základ 0%", "DPH - 12%", "DPH 21%", "Celkem bez DPH", "Celkem s DPH"]
        for col in num_cols:
            if col in df_final.columns:
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

        # 4. EXPORT
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb, ws = writer.book, writer.book.add_worksheet('Report')
            f_bold = wb.add_format({'bold': True, 'font_size': 12})
            f_num = wb.add_format({'num_format': '#,##0.00'})
            f_date = wb.add_format({'num_format': 'dd.mm.yyyy hh:mm:ss'})

            # Nadpis Tržba ze dne (dynamicky)
            d1 = st_range.strftime("%d.%m.")
            d2 = en_range.strftime("%d.%m. %Y")
            ws.write('B1', f"Tržba ze dne {d1} - {d2}", f_bold)

            # Zápis tabulky
            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)

            # Součty
            n = len(df_final)
            if n > 0:
                fr, lr, sr = 4, 4 + n - 1, 4 + n + 1
                # E = Forma úhrady, K = Celkem s DPH
                ws.write(sr, 2, "Hotovost celkem:", f_bold)
                ws.write_formula(sr, 3, f'=SUMIF(E{fr}:E{lr}, "*Hotově*", K{fr}:K{lr})', f_num)
                ws.write(sr+1, 2, "Kreditní karty celkem:", f_bold)
                ws.write_formula(sr+1, 3, f'=SUMIF(E{fr}:E{lr}, "*Kartou*", K{fr}:K{lr})', f_num)
            
            # Formátování šířky sloupců
            ws.set_column('A:K', 18)

        st.success(f"✅ Report pro období {d1} - {d2} vytvořen!")
        st.download_button("📥 Stáhnout hotový report", output.getvalue(), f"Uzaverka_{d1}_{d2}.xlsx")

    except Exception as e:
        st.error(f"Chyba: {e}")

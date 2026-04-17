import streamlit as st
import pandas as pd
import io
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")
st.title("🏨 Hotelový Reportér")

st.info("Nahrajte soubor .xlsx (uložený v Excelu jako Sešit Excel).")

uploaded_file = st.file_uploader("Nahrajte .xlsx soubor", type=["xlsx"])

if uploaded_file:
    try:
        # 1. NAČTENÍ - Načteme surová data bez jakýchkoliv hlaviček
        df = pd.read_excel(uploaded_file, header=None)

        # 2. DETEKCE POZIC (INDEXŮ) SLOUPCŮ
        header_row_idx = None
        col_indices = {}
        
        # Klíčová slova pro hledání sloupců
        targets = {
            "vystaveno": "vyst",
            "forma úhrady": "uhrad",
            "celkem s dph": "s dph",
            "stav": "stav",
            "číslo": "číslo",
            "základ 0%": "0%",
            "12%": "12%",
            "21%": "21%"
        }

        # Projdeme řádky a najdeme ten s hlavičkou
        for i in range(len(df)):
            row_as_list = [str(x).lower().strip() for x in df.iloc[i].values]
            if any("vystaveno" in str(x) for x in row_as_list):
                header_row_idx = i
                # Uložíme si pozice (čísla sloupců) pro všechny důležité údaje
                for target_key, search_str in targets.items():
                    for col_idx, cell_val in enumerate(row_as_list):
                        if search_str in cell_val:
                            col_indices[target_key] = col_idx
                break

        if header_row_idx is None or "vystaveno" not in col_indices:
            st.error("❌ Nepodařilo se identifikovat strukturu tabulky. Je v souboru sloupec 'Vystaveno'?")
            st.stop()

        # 3. ZPRACOVÁNÍ DAT (Práce s pozicemi .iloc)
        # Ořízneme data (vše pod nalezenou hlavičkou)
        df_data = df.iloc[header_row_idx + 1:].copy().reset_index(drop=True)

        # Vytvoříme pomocný sloupec s datem (pomocí nalezeného indexu sloupce Vystaveno)
        vyst_idx = col_indices["vystaveno"]
        df_data['temp_date'] = pd.to_datetime(df_data.iloc[:, vyst_idx], dayfirst=True, errors='coerce')
        
        # Vyhodíme řádky bez data (patičky, prázdné řádky)
        df_data = df_data.dropna(subset=['temp_date'])

        if df_data.empty:
            st.warning("⚠️ Žádná platná data k seřazení.")
            st.stop()

        # 4. FILTRACE 10:00 - 12:00
        min_d = df_data['temp_date'].min().date()
        st_range = datetime.combine(min_d, time(10, 0, 0))
        en_range = datetime.combine(min_d + pd.Timedelta(days=1), time(12, 0, 0))
        
        df_f = df_data[(df_data['temp_date'] >= st_range) & (df_data['temp_date'] <= en_range)].copy()

        # 5. SESTAVENÍ FINÁLNÍ TABULKY (přesné názvy pro export)
        # Mapujeme nalezené indexy na nové názvy
        export_data = {}
        pretty_names = {
            "vystaveno": "Vystaveno", "stav": "Stav", "číslo": "Číslo",
            "forma úhrady": "Forma úhrady", "základ 0%": "Základ 0%",
            "12%": "DPH 12%", "21%": "DPH 21%", "celkem s dph": "Celkem s DPH"
        }

        for key, idx in col_indices.items():
            export_data[pretty_names[key]] = df_f.iloc[:, idx]

        df_final = pd.DataFrame(export_data)

        # Čištění čísel (vše na numerické hodnoty)
        num_cols = ["Základ 0%", "DPH 12%", "DPH 21%", "Celkem s DPH"]
        for col in num_cols:
            if col in df_final.columns:
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

        # 6. EXPORT
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb, ws = writer.book, writer.book.add_worksheet('Report')
            f_bold, f_num = wb.add_format({'bold': True}), wb.add_format({'num_format': '#,##0.00'})
            
            d1, d2 = st_range.strftime("%d.%m."), en_range.strftime("%d.%m.%Y")
            ws.write('B1', f"Report tržeb {d1} - {d2}", f_bold)
            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)
            
            # SUMIF (výpočty na konci)
            n = len(df_final)
            if n > 0:
                fr, lr, sr = 4, 4 + n - 1, 4 + n + 1
                try:
                    cols = list(df_final.columns)
                    # Najdeme písmena sloupců v exportu
                    p_let = chr(65 + cols.index("Forma úhrady"))
                    t_let = chr(65 + cols.index("Celkem s DPH"))
                    ws.write(sr, 2, "Hotovost celkem:", f_bold)
                    ws.write_formula(sr, 3, f'=SUMIF({p_let}{fr}:{p_let}{lr}, "*Hotově*", {t_let}{fr}:{t_let}{lr})', f_num)
                    ws.write(sr + 1, 2, "Karty celkem:", f_bold)
                    ws.write_formula(sr + 1, 3, f'=SUMIF({p_let}{fr}:{p_let}{lr}, "*Kartou*", {t_let}{fr}:{t_let}{lr})', f_num)
                except: pass

        st.success("✅ Report úspěšně vygenerován!")
        st.download_button("📥 Stáhnout Excel", output.getvalue(), f"Uzaverka_{d1}{d2}.xlsx")

    except Exception as e:
        st.error(f"Kritická chyba: {e}")

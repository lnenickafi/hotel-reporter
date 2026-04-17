import streamlit as st
import pandas as pd
import io
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")
st.title("🏨 Hotelový Reportér")

with st.expander("💡 INSTRUKCE", expanded=True):
    st.write("Nahrajte soubor, který jste v Excelu uložili jako 'Sešit Excel (.xlsx)'.")

uploaded_file = st.file_uploader("Nahrajte .xlsx soubor", type=["xlsx"])

if uploaded_file:
    try:
        # 1. NAČTENÍ XLSX
        df = pd.read_excel(uploaded_file)

        # 2. VYČIŠTĚNÍ NÁZVŮ SLOUPCŮ (Zásadní krok!)
        # Odstraní mezery na začátku/konci a uvozovky ze všech názvů sloupců
        df.columns = [str(c).strip().replace('"', '') for c in df.columns]

        # 3. KONTROLA A PŘEVOD DATA
        # Hledáme sloupec Vystaveno (i kdyby se jmenoval trochu jinak)
        vyst_col = next((c for c in df.columns if "Vystaveno" in c), None)
        
        if not vyst_col:
            st.error(f"❌ Sloupec 'Vystaveno' nenalezen. Dostupné sloupce jsou: {', '.join(df.columns[:5])}...")
            st.stop()

        # Převod na datum - bereme Series (1D pole), abychom se vyhnuli chybě 'arg must be a list'
        # Použijeme .squeeze(), aby to byl vždy jeden sloupec
        date_series = df[vyst_col]
        if isinstance(date_series, pd.DataFrame):
            date_series = date_series.iloc[:, 0]

        df['dt_fixed'] = pd.to_datetime(date_series, dayfirst=True, errors='coerce')
        
        # Odstranění řádků bez data (patičky, prázdné řádky)
        df = df.dropna(subset=['dt_fixed'])

        if df.empty:
            st.warning("⚠️ V souboru nebyla nalezena žádná platná data k seřazení.")
            st.stop()

        # 4. FILTRACE ČASU (10:00 - 12:00 druhý den)
        min_date = df['dt_fixed'].min().date()
        start_range = datetime.combine(min_date, time(10, 0, 0))
        end_range = datetime.combine(min_date + pd.Timedelta(days=1), time(12, 0, 0))
        
        df_f = df[(df['dt_fixed'] >= start_range) & (df['dt_fixed'] <= end_range)].copy()

        # 5. MAPOVÁNÍ SLOUPCŮ
        mapping = {
            "Vystaveno": "Vystaveno",
            "Stav": "Stav",
            "Číslo": "Číslo",
            "Forma úhrady": "Forma úhrady",
            "Základ 0%": "Základ 0%",
            "12%": "DPH 12%",
            "21%": "DPH 21%",
            "Celkem s DPH": "Celkem s DPH"
        }
        
        # Najdeme reálné sloupce v souboru podle klíčových slov
        final_cols_map = {}
        for real_c in df_f.columns:
            for key, target in mapping.items():
                if key.lower() in str(real_c).lower():
                    final_cols_map[real_c] = target
                    break
        
        df_final = df_f[list(final_cols_map.keys())].rename(columns=final_cols_map)

        # Čištění čísel (v XLSX už bývají čísla, ale pro jistotu)
        for col in df_final.columns:
            if any(x in col for x in ["Základ", "DPH", "Celkem"]):
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

        # 6. EXPORT DO EXCELU
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb, ws = writer.book, writer.book.add_worksheet('Report')
            f_bold, f_num = wb.add_format({'bold': True}), wb.add_format({'num_format': '#,##0.00'})
            
            d1, d2 = start_range.strftime("%d.%m."), end_range.strftime("%d.%m.%Y")
            ws.write('B1', f"Tržba {d1} - {d2}", f_bold)
            df_final.to_excel(writer, sheet_name='Report', index=False, startrow=2)
            
            # Dynamické součty SUMIF
            n = len(df_final)
            if n > 0:
                fr, lr, sr = 4, 4 + n - 1, 4 + n + 1
                try:
                    c_idx = list(df_final.columns)
                    # Najdeme písmena sloupců pro vzorec
                    p_col = chr(65 + c_idx.index("Forma úhrady"))
                    t_col = chr(65 + c_idx.index("Celkem s DPH"))
                    
                    ws.write(sr, 2, "Hotovost celkem:", f_bold)
                    ws.write_formula(sr, 3, f'=SUMIF({p_col}{fr}:{p_col}{lr}, "*Hotově*", {t_col}{fr}:{t_col}{lr})', f_num)
                    ws.write(sr + 1, 2, "Karty celkem:", f_bold)
                    ws.write_formula(sr + 1, 3, f'=SUMIF({p_col}{fr}:{p_col}{lr}, "*Kartou*", {t_col}{fr}:{t_col}{lr})', f_num)
                except:
                    pass

        st.success("✅ Report úspěšně vytvořen!")
        st.download_button(
            label="📥 Stáhnout upravený Excel",
            data=output.getvalue(),
            file_name=f"Report_{d1}{d2}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Chyba při zpracování: {e}")

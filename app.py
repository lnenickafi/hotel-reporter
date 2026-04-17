import streamlit as st
import pandas as pd
import io
from datetime import datetime, time

st.set_page_config(page_title="Hotelový Reportér", page_icon="🏨")
st.title("🏨 Hotelový Reportér")

with st.expander("💡 POSLEDNÍ KROK K ÚSPĚCHU", expanded=True):
    st.write("Nahrajte soubor .xlsx (uložený v Excelu jako Sešit Excel).")

uploaded_file = st.file_uploader("Nahrajte .xlsx soubor", type=["xlsx"])

if uploaded_file:
    try:
        # 1. NAČTENÍ XLSX - Načteme surová data bez hlavičky
        df = pd.read_excel(uploaded_file, header=None)

        # 2. AUTOMATICKÁ DETEKCE HLAVIČKY A SLOUPCŮ
        # Hledáme řádek, kde je slovo "Vystaveno"
        header_row_idx = None
        vystaveno_col_idx = None
        
        for i in range(len(df)):
            row = [str(x).strip() for x in df.iloc[i].values]
            for idx, val in enumerate(row):
                if "Vystaveno" in val:
                    header_row_idx = i
                    vystaveno_col_idx = idx
                    break
            if header_row_idx is not None:
                break
        
        if header_row_idx is None:
            st.error("❌ V souboru nebyl nalezen sloupec 'Vystaveno'.")
            st.stop()

        # Nastavíme názvy sloupců z nalezeného řádku a ořízneme data
        df.columns = [str(c).strip() for c in df.iloc[header_row_idx]]
        df = df.iloc[header_row_idx + 1:].reset_index(drop=True)

        # 3. PŘEVOD DATA - Tady opravujeme chybu 'arg must be a list'
        # Použijeme iloc a index, abychom měli 100% jistotu, že bereme 1D pole
        raw_dates = df.iloc[:, vystaveno_col_idx]
        
        # Převedeme na datum (errors='coerce' vymaže neplatné řádky/patičky)
        df['dt_fixed'] = pd.to_datetime(raw_dates, dayfirst=True, errors='coerce')
        df = df.dropna(subset=['dt_fixed'])

        if df.empty:
            st.warning("⚠️ Žádná platná data k seřazení (zkontrolujte formát data).")
            st.stop()

        # 4. FILTRACE ČASU (10:00 - 12:00 druhý den)
        min_date = df['dt_fixed'].min().date()
        start_range = datetime.combine(min_date, time(10, 0, 0))
        end_range = datetime.combine(min_date + pd.Timedelta(days=1), time(12, 0, 0))
        
        df_f = df[(df['dt_fixed'] >= start_range) & (df['dt_fixed'] <= end_range)].copy()

        # 5. DYNAMICKÝ VÝBĚR SLOUPCŮ PRO FINÁLNÍ TABULKU
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
        
        final_cols = []
        for key, target in mapping.items():
            for col_name in df_f.columns:
                if key.lower() in str(col_name).lower():
                    df_f = df_f.rename(columns={col_name: target})
                    final_cols.append(target)
                    break
        
        # Odstraníme případné duplicity v názvech, které by dělaly neplechu
        final_df = df_f[list(dict.fromkeys(final_cols))].copy()

        # Čištění čísel
        for col in final_df.columns:
            if any(x in col for x in ["Základ", "DPH", "Celkem"]):
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce').fillna(0)

        # 6. EXPORT DO EXCELU
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb, ws = writer.book, writer.book.add_worksheet('Report')
            f_bold, f_num = wb.add_format({'bold': True}), wb.add_format({'num_format': '#,##0.00'})
            
            d1, d2 = start_range.strftime("%d.%m."), end_range.strftime("%d.%m.%Y")
            ws.write('B1', f"Tržba {d1} - {d2}", f_bold)
            final_df.to_excel(writer, sheet_name='Report', index=False, startrow=2)
            
            # SUMIF vzorce
            n = len(final_df)
            if n > 0:
                fr, lr, sr = 4, 4 + n - 1, 4 + n + 1
                try:
                    cols = list(final_df.columns)
                    p_letter = chr(65 + cols.index("Forma úhrady"))
                    t_letter = chr(65 + cols.index("Celkem s DPH"))
                    ws.write(sr, 2, "Hotovost celkem:", f_bold)
                    ws.write_formula(sr, 3, f'=SUMIF({p_letter}{fr}:{p_letter}{lr}, "*Hotově*", {t_letter}{fr}:{t_letter}{lr})', f_num)
                    ws.write(sr + 1, 2, "Karty celkem:", f_bold)
                    ws.write_formula(sr + 1, 3, f'=SUMIF({p_letter}{fr}:{p_letter}{lr}, "*Kartou*", {t_letter}{fr}:{t_letter}{lr})', f_num)
                except: pass

        st.success("✅ Report úspěšně vytvořen!")
        st.download_button("📥 Stáhnout Excel", output.getvalue(), f"Uzaverka_{d1}{d2}.xlsx")

    except Exception as e:
        st.error(f"Kritická chyba: {e}")

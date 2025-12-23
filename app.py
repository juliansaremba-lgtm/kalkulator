import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Tectum Group - Kalkulator", layout="wide")

st.title("🏗️ Kalkulator Rozliczeń Podwykonawcy")
st.info("Wypełnij tabelę poniżej, a następnie kliknij przycisk na dole, aby pobrać gotowy Excel z formułami.")

# Lista materiałów do wyboru
mozliwe_operacje = [
    "Grunt bitumiczny Sopradere",
    "Papa podkładowa Sopralene Flam 180",
    "Papa nawierzchniowa Sopralene Flam Jardin",
    "Folia PE gr. 0,2mm",
    "Ravatherm XPS 300SL gr. 8cm",
    "Ravatherm XPS 500SL gr. 5cm",
    "Mata dyfuzyjna Delta Vent RR",
    "Drenaż Floraxx 60H+",
    "Geowłóknina 105g/m2 Polyfelt TS10",
    "Wywinięcie na ściany budynku",
    "Wywinięcie na oczep",
    "Dylatacja",
    "Wpusty - izolacja"
]

# Tabela edytowalna
if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame(
        [[f"Lp.{i+1}", "Grunt bitumiczny Sopradere", "m2", 0.0, 0.0] for i in range(15)],
        columns=["Nr", "Opis", "Jedn.", "Ilość", "Stawka"]
    )

edited_df = st.data_editor(
    st.session_state.df,
    column_config={
        "Opis": st.column_config.SelectboxColumn("Opis prac", options=mozliwe_operacje, width="large"),
        "Jedn.": st.column_config.SelectboxColumn("Jedn.", options=["m2", "mb", "szt.", "kpl."]),
        "Ilość": st.column_config.NumberColumn("Ilość", min_value=0.0, format="%.2f"),
        "Stawka": st.column_config.NumberColumn("Stawka [PLN]", min_value=0.0, format="%.2f"),
    },
    num_rows="dynamic",
    use_container_width=True
)

# Przycisk generowania
if st.button("📥 Generuj plik Excel do rozliczeń"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        edited_df.to_excel(writer, index=False, sheet_name='Rozliczenie')
        workbook = writer.book
        worksheet = writer.sheets['Rozliczenie']
        
        # Formaty
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#CFE2F3', 'border': 1})
        num_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        pct_fmt = workbook.add_format({'num_format': '0%', 'border': 1, 'bg_color': '#FFF2CC'})
        
        # Dodanie kolumn miesięcznych w wygenerowanym Excelu
        last_col = 6 + 4*2
        for m in range(1, 5):
            col_idx = 6 + (m-1)*2
            worksheet.write(0, col_idx, f"Miesiąc {m} [%]", header_fmt)
            worksheet.write(0, col_idx+1, f"Miesiąc {m} [PLN]", header_fmt)
            for row in range(1, len(edited_df) + 1):
                worksheet.write(row, col_idx, 0, pct_fmt)
                worksheet.write_formula(row, col_idx+1, f'={chr(65+col_idx)}{row+1}*F{row+1}', num_fmt)
        
        for row in range(1, len(edited_df) + 1):
            worksheet.write_formula(row, 5, f'=D{row+1}*E{row+1}', num_fmt)

    st.download_button(
        label="✅ Pobierz plik teraz",
        data=output.getvalue(),
        file_name="Rozliczenie_Podwykonawcy.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

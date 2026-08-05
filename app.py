import streamlit as st
import pandas as pd
import re
from pypdf import PdfReader
import io

st.set_page_config(page_title="Stücklisten Extraktor", layout="centered")

st.title("Stromlaufplan Stücklisten-Extraktor")
st.write("Lade deinen Stromlaufplan als PDF hoch, um automatisch eine konsolidierte Excel-Datei zu erstellen.")

uploaded_file = st.file_uploader("PDF-Datei auswählen", type=["pdf"])
keyword = st.text_input("Filter-Schlüsselwort (optional)", placeholder="z. B. Siemens, Phoenix, Typ")

if uploaded_file is not None:
    if st.button("Stückliste als Excel generieren"):
        reader = PdfReader(uploaded_file)
        extracted_rows = []

        for page in reader.pages:
            text = page.extract_text()
            if not text:
                continue
            
            lines = text.split('\n')
            for line in lines:
                line_str = line.strip()
                if not line_str:
                    continue

                if keyword and keyword.lower() not in line_str.lower():
                    continue

                parts = re.split(r'\s{2,}|\t', line_str)
                if len(parts) >= 2:
                    extracted_rows.append(parts)

        if extracted_rows:
            df = pd.DataFrame(extracted_rows)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="Stückliste")
            
            st.success("Erfolgreich konvertiert!")
            st.download_button(
                label="Excel-Datei herunterladen",
                data=output.getvalue(),
                file_name="Konsolidierte_Stueckliste.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Keine passenden Zeilen in der PDF gefunden.")

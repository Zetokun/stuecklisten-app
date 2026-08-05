import io
import re
import pandas as pd
from pypdf import PdfReader
import streamlit as st

st.set_page_config(page_title="Stücklisten Extraktor", page_icon="📊")

st.title("📊 Stromlaufplan Stücklisten-Extraktor")
st.write("Laden Sie Ihre PDF-Datei hoch, um die Stückliste nach Excel zu exportieren.")

uploaded_file = st.file_uploader("PDF-Datei auswählen", type=["pdf"])

if uploaded_file is not None:
    try:
        reader = PdfReader(uploaded_file)
        rows = []
        
        for page in reader.pages:
            text = page.extract_text()
            if text:
                for line in text.split('\n'):
                    line_str = line.strip()
                    if line_str:
                        parts = re.split(r'\s{2,}|\t', line_str)
                        if len(parts) >= 2:
                            rows.append(parts)

        if not rows:
            st.warning("Keine Tabellendaten im PDF gefunden.")
        else:
            df = pd.DataFrame(rows)
            
            # Excel im Speicher erstellen
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, header=False, sheet_name='Stückliste')
            
            excel_data = output.getvalue()

            st.success("Erfolgreich extrahiert!")
            st.download_button(
                label="📥 Excel-Datei herunterladen",
                data=excel_data,
                file_name="Stueckliste.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Fehler bei der PDF-Verarbeitung: {e}")

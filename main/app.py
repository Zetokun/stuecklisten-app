import io
import re
import pandas as pd
from pypdf import PdfReader
import streamlit as st

# Seiteneinstellungen
st.set_page_config(page_title="Stücklisten Extraktor", page_icon="📊", layout="wide")

st.title("📊 Stücklisten PDF-Extraktor")
st.write("Lade deine Stücklisten-PDF hoch, um die Daten als Excel-Datei zu exportieren.")

# Datei-Uploader für PDFs
uploaded_file = st.file_uploader("PDF-Datei auswählen", type=["pdf"])

if uploaded_file is not None:
    try:
        reader = PdfReader(uploaded_file)
        rows = []
        
        # Jede Seite des PDFs durchgehen
        for page in reader.pages:
            text = page.extract_text()
            if text:
                for line in text.split('\n'):
                    line_str = line.strip()
                    if line_str:
                        # Trennung nach Spalten (bei mehreren Leerzeichen oder Tabs)
                        parts = re.split(r'\s{2,}|\t', line_str)
                        if len(parts) >= 1:
                            rows.append(parts)

        if not rows:
            st.warning("Keine lesbaren Textdaten in der PDF gefunden.")
        else:
            # Erstelle DataFrame aus den ausgelesenen Zeilen
            df = pd.DataFrame(rows)
            
            st.subheader("Vorschau der extrahierten Daten:")
            st.dataframe(df.head(25), use_container_width=True)
            
            # Excel-Datei im Arbeitsspeicher erstellen
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, header=False, sheet_name='Stückliste')
            
            excel_data = output.getvalue()

            st.success("PDF erfolgreich verarbeitet!")
            
            # Download-Button anzeigen
            st.download_button(
                label="📥 Excel-Datei herunterladen (.xlsx)",
                data=excel_data,
                file_name="Stueckliste.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Fehler bei der Verarbeitung der PDF-Datei: {e}")

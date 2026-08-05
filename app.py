import io
import re
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pypdf import PdfReader
import streamlit as st

st.set_page_config(page_title="Stücklisten Extraktor", page_icon="📊")

st.title("📊 Stromlaufplan Stücklisten-Extraktor")
st.write("Laden Sie Ihre PDF-Datei hoch, um automatisch eine Excel-Stückliste zu erstellen.")

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
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Stückliste"
            ws.views.sheetView[0].showGridLines = True

            for r_idx, row in enumerate(rows, start=1):
                for c_idx, val in enumerate(row, start=1):
                    ws.cell(row=r_idx, column=c_idx, value=val)

            wb.save(output)
            output.seek(0)

            st.success("Erfolgreich extrahiert!")
            st.download_button(
                label="📥 Excel-Datei herunterladen",
                data=output,
                file_name="Stueckliste.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Fehler bei der PDF-Verarbeitung: {e}")

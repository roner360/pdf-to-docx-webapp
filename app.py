import io
import os
import tempfile
import streamlit as st
from pdf2docx import Converter

st.set_page_config(page_title="PDF → DOCX", page_icon="📄")
st.title("📄 PDF → DOCX Converter (layout + immagini)")

uploaded = st.file_uploader("Carica un PDF", type=["pdf"])

col1, col2 = st.columns(2)
with col1:
    start_page = st.number_input("Pagina iniziale (0 = prima)", min_value=0, value=0, step=1)
with col2:
    end_page = st.number_input("Pagina finale (None = tutte)", min_value=0, value=0, step=1)

use_all_pages = st.checkbox("Converti tutte le pagine", value=True)

if uploaded:
    st.info("Conversione in corso… (può richiedere un po’ per PDF lunghi)")

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = os.path.join(tmpdir, "input.pdf")
            docx_path = os.path.join(tmpdir, "output.docx")

            # salva upload su file temporaneo
            with open(pdf_path, "wb") as f:
                f.write(uploaded.read())

            # converti
            cv = Converter(pdf_path)
            if use_all_pages:
                cv.convert(docx_path)  # tutte le pagine
            else:
                # pdf2docx usa start/end (end esclusivo in alcune versioni): qui usiamo end_page+1 per includerla
                cv.convert(docx_path, start=int(start_page), end=int(end_page) + 1)
            cv.close()

            # leggi docx e offri download
            with open(docx_path, "rb") as f:
                docx_bytes = f.read()

        st.success("Fatto ✅")
        st.download_button(
            "Scarica DOCX",
            data=docx_bytes,
            file_name=uploaded.name.rsplit(".", 1)[0] + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        st.caption("Nota: conversioni perfette al 100% non esistono; dipende da com’è fatto il PDF.")
    except Exception as e:
        st.error(f"Errore di conversione: {e}")
        st.write("Se il PDF è una scansione (immagine), serve OCR: il DOCX sarà comunque limitato.")

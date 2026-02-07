import io
import streamlit as st
import pdfplumber
from docx import Document

st.set_page_config(page_title="PDF → DOCX Converter", page_icon="📄")

st.title("📄 PDF → DOCX Converter")
st.write("Upload a text-based PDF and download a DOCX.")

uploaded = st.file_uploader("Upload PDF", type=["pdf"])

def pdf_to_docx_bytes(pdf_bytes: bytes) -> bytes:
    pdf_file = io.BytesIO(pdf_bytes)
    doc = Document()

    with pdfplumber.open(pdf_file) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            doc.add_heading(f"Page {i}", level=2)
            for line in text.splitlines():
                doc.add_paragraph(line)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read()

if uploaded:
    try:
        st.info("Converting…")
        docx_data = pdf_to_docx_bytes(uploaded.read())

        st.success("Done ✅")
        st.download_button(
            "Download DOCX",
            data=docx_data,
            file_name=uploaded.name.rsplit(".", 1)[0] + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        st.error(f"Conversion failed: {e}")
        st.write("If the PDF is scanned (image-only), you’ll need OCR.")

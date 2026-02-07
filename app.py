import io, re, os, tempfile
import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn

st.set_page_config(page_title="PDF → DOCX (layout)", page_icon="📄")
st.title("📄 PDF → DOCX (mantieni struttura lista con immagini)")

uploaded = st.file_uploader("Carica un PDF", type=["pdf"])

def normalize_text(t: str) -> str:
    t = t.replace("\u2060", "").strip()
    lines = [ln.strip() for ln in t.splitlines() if ln.strip()]
    code_lines = [ln for ln in lines if re.search(r"\(X[0-9A-Z]+\)", ln)]
    other = [ln for ln in lines if ln not in code_lines]
    return "\n".join(other + code_lines)

def extract_rows(page):
    blocks = page.get_text("dict")["blocks"]

    # prendi le immagini “vere” (non bordi), tipicamente più pesanti
    img_blocks = [b for b in blocks if b["type"] == 1 and b.get("size", 0) > 10000]
    img_blocks = sorted(img_blocks, key=lambda b: b["bbox"][1])

    text_blocks = [b for b in blocks if b["type"] == 0]

    # costruisci bande verticali per riga (in base alla distanza tra immagini)
    rows = []
    for i, img in enumerate(img_blocks):
        y0 = img["bbox"][1]
        y1 = img["bbox"][3]
        next_y0 = img_blocks[i + 1]["bbox"][1] if i + 1 < len(img_blocks) else page.rect.height + 1

        top = (img_blocks[i - 1]["bbox"][3] + y0) / 2 if i > 0 else y0 - 5
        bottom = (y1 + next_y0) / 2

        # testo a destra dell’immagine (x > ~105) che cade nella banda verticale
        parts = []
        for tb in text_blocks:
            x0, ty0, x1, ty1 = tb["bbox"]
            if x0 < 105:
                continue
            if ty1 < top or ty0 > bottom:
                continue

            s = ""
            for line in tb["lines"]:
                for span in line["spans"]:
                    s += span["text"]
                s += "\n"
            s = s.strip()
            if s:
                parts.append((tb["bbox"][1], s))

        parts.sort(key=lambda x: x[0])
        full_text = normalize_text("\n".join([p for _, p in parts]))

        rows.append({"image_bytes": img["image"], "text": full_text})

    return rows

def remove_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "nil")
        borders.append(e)
    tblPr.append(borders)

def convert_pdf_to_docx(pdf_bytes: bytes) -> bytes:
    doc = Document()
    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")

    # header dal top della prima pagina
    first = pdf[0]
    header_blocks = [b for b in first.get_text("dict")["blocks"] if b["type"] == 0 and b["bbox"][1] < 190]
    header_blocks.sort(key=lambda b: b["bbox"][1])

    for idx, b in enumerate(header_blocks):
        t = ""
        for line in b["lines"]:
            for span in line["spans"]:
                t += span["text"]
            t += "\n"
        t = t.replace("\u2060", "").strip()
        if not t:
            continue
        p = doc.add_paragraph(t)
        if idx == 0:
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(22)
        elif idx == 1:
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(16)
        else:
            for r in p.runs:
                r.font.size = Pt(11)

    doc.add_paragraph("")

    # righe prodotto come tabella 2 colonne (immagine + testo)
    for page in pdf:
        for row in extract_rows(page):
            table = doc.add_table(rows=1, cols=2)
            table.alignment = WD_TABLE_ALIGNMENT.LEFT
            table.autofit = False
            table.columns[0].width = Inches(1.2)
            table.columns[1].width = Inches(5.5)
            remove_table_borders(table)

            cell_img = table.cell(0, 0)
            p = cell_img.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(io.BytesIO(row["image_bytes"]), width=Inches(0.9))

            cell_txt = table.cell(0, 1)
            p2 = cell_txt.paragraphs[0]
            lines = row["text"].splitlines() if row["text"] else []
            for i, line in enumerate(lines):
                run = p2.add_run(line)
                run.font.size = Pt(11)
                if i < len(lines) - 1:
                    p2.add_run("\n")

            doc.add_paragraph("")

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read()

if uploaded:
    try:
        st.info("Converto…")
        docx_bytes = convert_pdf_to_docx(uploaded.read())
        st.success("Fatto ✅")
        st.download_button(
            "Scarica DOCX",
            data=docx_bytes,
            file_name=uploaded.name.rsplit(".", 1)[0] + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        st.error(f"Errore: {e}")

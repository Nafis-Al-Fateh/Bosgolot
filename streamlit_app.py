import streamlit as st
import fitz
import pytesseract
import cv2
import numpy as np
import os
import tempfile
import shutil
from PIL import Image
from docx import Document
from docx.shared import Inches
from openpyxl import Workbook
from sklearn.cluster import KMeans
import camelot

st.set_page_config(page_title="Advanced PDF OCR System", layout="wide")

st.title("📄 Advanced PDF → DOCX + Excel")
st.markdown("Supports Bangla + English | Multi-column | Images | Tables")

# -----------------------------
# IMAGE PREPROCESSING
# -----------------------------
def preprocess_image(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (5,5), 0)
    thresh = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 11, 2
    )
    return thresh


# -----------------------------
# OCR FUNCTION
# -----------------------------
def ocr_page(page):
    pix = page.get_pixmap(dpi=300)
    img = np.frombuffer(pix.samples, dtype=np.uint8)
    img = img.reshape(pix.height, pix.width, pix.n)

    processed = preprocess_image(img)

    text = pytesseract.image_to_string(
        processed,
        lang="eng+ben",
        config="--oem 3 --psm 6"
    )

    return [{"type": "text", "content": text, "y": 0}]


# -----------------------------
# MULTI-COLUMN PARSER
# -----------------------------
def parse_layout(page):
    blocks = page.get_text("blocks")
    if not blocks:
        return []

    xs = np.array([[b[0]] for b in blocks])

    if len(xs) > 1:
        try:
            kmeans = KMeans(n_clusters=2, random_state=0).fit(xs)
            labels = kmeans.labels_

            col1 = [blocks[i] for i in range(len(blocks)) if labels[i] == 0]
            col2 = [blocks[i] for i in range(len(blocks)) if labels[i] == 1]

            col1.sort(key=lambda x: x[1])
            col2.sort(key=lambda x: x[1])

            ordered = col1 + col2
        except:
            ordered = sorted(blocks, key=lambda x: x[1])
    else:
        ordered = blocks

    parsed = []
    for b in ordered:
        parsed.append({
            "type": "text",
            "content": b[4],
            "y": b[1]
        })

    return parsed


# -----------------------------
# IMAGE EXTRACTION
# -----------------------------
def extract_images(page, temp_dir):
    images = []
    image_list = page.get_images(full=True)

    for img_index, img in enumerate(image_list):
        xref = img[0]
        base_image = page.parent.extract_image(xref)
        image_bytes = base_image["image"]

        image_path = os.path.join(temp_dir, f"img_{xref}.png")
        with open(image_path, "wb") as f:
            f.write(image_bytes)

        images.append(image_path)

    return images


# -----------------------------
# TABLE EXTRACTION
# -----------------------------
def extract_tables(pdf_path, page_number):
    try:
        tables = camelot.read_pdf(
            pdf_path,
            pages=str(page_number + 1),
            flavor="stream"
        )
        return [t.df.values.tolist() for t in tables]
    except:
        return []


# -----------------------------
# MAIN PROCESSING
# -----------------------------
def process_pdf(pdf_path):

    doc = fitz.open(pdf_path)
    document = Document()
    workbook = Workbook()
    workbook.remove(workbook.active)

    temp_dir = tempfile.mkdtemp()

    for page_num in range(len(doc)):
        page = doc[page_num]

        text = page.get_text().strip()

        # ---------- TEXT EXTRACTION ----------
        if len(text) > 20:
            layout_data = parse_layout(page)
        else:
            layout_data = ocr_page(page)

        # ---------- ADD TEXT TO DOC ----------
        for block in layout_data:
            document.add_paragraph(block["content"])

        # ---------- ADD IMAGES ----------
        images = extract_images(page, temp_dir)
        for img_path in images:
            document.add_picture(img_path, width=Inches(4))

        document.add_page_break()

        # ---------- TABLES ----------
        tables = extract_tables(pdf_path, page_num)

        sheet = workbook.create_sheet(f"Page_{page_num+1}")
        for table in tables:
            for row in table:
                sheet.append(row)
            sheet.append([])

    # Save DOCX
    doc_output = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    document.save(doc_output.name)

    # Save Excel
    excel_output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    workbook.save(excel_output.name)

    shutil.rmtree(temp_dir)

    return doc_output.name, excel_output.name


# -----------------------------
# STREAMLIT UI
# -----------------------------
uploaded_file = st.file_uploader("Upload PDF (Max 10MB)", type=["pdf"])

if uploaded_file:

    if uploaded_file.size > 10 * 1024 * 1024:
        st.error("File exceeds 10MB limit.")
    else:
        temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        temp_pdf.write(uploaded_file.read())
        temp_pdf.close()

        if st.button("Process PDF"):
            try:
                with st.spinner("Processing..."):
                    doc_path, excel_path = process_pdf(temp_pdf.name)

                st.success("Processing Complete!")

                with open(doc_path, "rb") as f:
                    st.download_button(
                        "Download DOCX",
                        f,
                        file_name="output.docx"
                    )

                with open(excel_path, "rb") as f:
                    st.download_button(
                        "Download Excel",
                        f,
                        file_name="output.xlsx"
                    )

                # Auto cleanup
                os.remove(temp_pdf.name)
                os.remove(doc_path)
                os.remove(excel_path)

            except Exception as e:
                st.error("Processing failed. Please restart and try again.")
                st.exception(e)

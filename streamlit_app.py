import streamlit as st
import fitz
import pytesseract
import cv2
import numpy as np
import os
import tempfile
import shutil
from docx import Document
from docx.shared import Inches
from openpyxl import Workbook
from sklearn.cluster import KMeans
import camelot

st.set_page_config(page_title="Advanced Bangla PDF Converter", layout="wide")

st.title("📄 Advanced Bangla PDF → DOCX + Excel")
st.markdown("Supports Bijoy ANSI + Unicode + OCR fallback")

# -----------------------------
# BIJOY → UNICODE MAPPING (Core Characters)
# -----------------------------
def bijoy_to_unicode(text):

    mapping = {
        "Av": "আ",
        "A": "অ",
        "B": "ই",
        "C": "ঈ",
        "D": "উ",
        "E": "ঊ",
        "F": "ঋ",
        "G": "এ",
        "H": "ঐ",
        "I": "ও",
        "J": "ঔ",
        "K": "ক",
        "L": "খ",
        "M": "গ",
        "N": "ঘ",
        "O": "ঙ",
        "P": "চ",
        "Q": "ছ",
        "R": "জ",
        "S": "ঝ",
        "T": "ঞ",
        "U": "ট",
        "V": "ঠ",
        "W": "ড",
        "X": "ঢ",
        "Y": "ণ",
        "Z": "ত",
        "a": "থ",
        "b": "দ",
        "c": "ধ",
        "d": "ন",
        "e": "প",
        "f": "ফ",
        "g": "ব",
        "h": "ভ",
        "i": "ম",
        "j": "য",
        "k": "র",
        "l": "ল",
        "m": "শ",
        "n": "ষ",
        "o": "স",
        "p": "হ",
        "q": "ড়",
        "r": "ঢ়",
        "s": "য়",
        "t": "ং",
        "u": "ঃ",
        "v": "ঁ"
    }

    for key, value in mapping.items():
        text = text.replace(key, value)

    return text


# -----------------------------
# DETECT BIJOY TEXT
# -----------------------------
def looks_like_bijoy(text):
    suspicious_patterns = ["wj", "‡", "†", "›"]
    for p in suspicious_patterns:
        if p in text:
            return True
    return False


# -----------------------------
# OCR FALLBACK
# -----------------------------
def ocr_page(page):

    pix = page.get_pixmap(dpi=400)
    img = np.frombuffer(pix.samples, dtype=np.uint8)
    img = img.reshape(pix.height, pix.width, pix.n)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.equalizeHist(gray)

    thresh = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 15, 10
    )

    text = pytesseract.image_to_string(
        thresh,
        lang="ben+eng",
        config="--oem 3 --psm 6"
    )

    return text


# -----------------------------
# MULTI COLUMN PARSER
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

    return ordered


# -----------------------------
# MAIN PROCESS
# -----------------------------
def process_pdf(pdf_path):

    doc = fitz.open(pdf_path)
    document = Document()
    workbook = Workbook()
    workbook.remove(workbook.active)

    for page_num in range(len(doc)):
        page = doc[page_num]

        raw_text = page.get_text()

        # -------- BIJOY DETECTION --------
        if looks_like_bijoy(raw_text):
            text = bijoy_to_unicode(raw_text)
        elif len(raw_text.strip()) > 20:
            blocks = parse_layout(page)
            text = ""
            for b in blocks:
                text += b[4] + "\n"
        else:
            text = ocr_page(page)

        # -------- DOCX --------
        document.add_paragraph(text)
        document.add_page_break()

        # -------- EXCEL --------
        sheet = workbook.create_sheet(f"Page_{page_num+1}")
        for line in text.split("\n"):
            sheet.append([line])

    # Save files
    doc_output = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    document.save(doc_output.name)

    excel_output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    workbook.save(excel_output.name)

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
                    st.download_button("Download DOCX", f, file_name="output.docx")

                with open(excel_path, "rb") as f:
                    st.download_button("Download Excel", f, file_name="output.xlsx")

                os.remove(temp_pdf.name)
                os.remove(doc_path)
                os.remove(excel_path)

            except Exception as e:
                st.error("Processing failed. Please restart and try again.")
                st.exception(e)

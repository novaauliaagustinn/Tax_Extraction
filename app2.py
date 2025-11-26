try:
    import fitz  # Lokal
except ImportError:
    import pymupdf as fitz  # Streamlit Cloud

import streamlit as st
import pandas as pd
from io import BytesIO
import re



# ============================
#  CONFIG TEMA & LAYOUT
# ============================
st.set_page_config(
    page_title="Tax Extraction",
    layout="centered"
)

# ============================
# SEMBUNYIKAN PREVIEW FILE
# ============================
st.markdown("""
<style>
.uploadedFile {display: none !important;}
.stUploadedFile {display: none !important;}
</style>
""", unsafe_allow_html=True)

# ============================
# BALIK URUTAN TAMPILAN FILE
# ============================
st.markdown("""
<style>
.stFileUploader > div:nth-child(1) > div {
    display: flex;
    flex-direction: column-reverse;
}
</style>
""", unsafe_allow_html=True)

# ============================
#  PERBESAR LIST FILE (MAX 10)
# ============================
st.markdown("""
<style>
/* Tinggikan container list file agar muat 10 file */
.stFileUploader > div > div {
    max-height: 450px !important; /* muat sekitar 10 file */
    overflow-y: auto !important;
}
</style>
""", unsafe_allow_html=True)


# CSS agar width max 750px
st.markdown("""
<style>
.main {
    max-width: 750px;
    margin: auto;
    padding-top: 20px;
}
</style>
""", unsafe_allow_html=True)

# ============================
#  HEADER
# ============================
st.markdown("""
<div style='text-align:left; padding: 10px 0;'>
    <h1 style='margin-bottom: 0; font-size: 40px; color:#150F3D;'>
        Tax Extraction
    </h1>
    <p style='font-size:16px; color:#555;'>This tool uses Artificial Intelligence–powered Natural Language Processing to accurately extract tax data from PDF documents by combining predefined coordinate-based regions with advanced Regular Expression (Regex) matching for more robust and precise data extraction</p>
    <hr style='margin-top:15px;'>
</div>
""", unsafe_allow_html=True)



# ============================
#  UPLOAD PDF
# ============================
uploaded_files = st.file_uploader("Upload PDF", type=["pdf"], accept_multiple_files=True)

# 🔥 Urutkan file sesuai tampilan (paling baru di atas)
if uploaded_files:
    uploaded_files = list(uploaded_files)[::-1]


# ===================================================
#  FUNGSI EKSTRAKSI XY
# ===================================================
def extract_by_xy(page, x0, y0, x1, y1, pdf_w, pdf_h, png_w, png_h):
    scale_x = pdf_w / png_w
    scale_y = pdf_h / png_h
    rect = fitz.Rect(
        x0 * scale_x, y0 * scale_y,
        x1 * scale_x, y1 * scale_y
    )
    return page.get_text("text", clip=rect).strip()

# ================================
# FUNGSI: CARI NOMOR DOKUMEN
# ================================
def find_nomor_dokumen(text):
    pattern = r"Nomor\s*Dokumen\s*[:\-]\s*([^\n\r]+)"
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else ""

# ================================
# FUNGSI: CARI NAMA PEMOTONG / PEMUNGUT
# ================================
def find_nama_pemungut(text):
    header = re.search(
        r"NAMA\s*PEMOTONG.*?PEMUNGUT\s*\n\s*PPh\s*[:]*",
        text,
        re.IGNORECASE | re.DOTALL
    )
    if not header:
        return ""

    start = header.end()
    after = text[start:].strip().split("\n")

    result_lines = []
    for line in after:
        line = line.strip()

        if line == "":
            break
        if re.search(r"C\.4", line, re.IGNORECASE):
            break
        if re.match(r"^\s*(Nomor|No|DPP|Jenis|Tarif|PPh)", line, re.IGNORECASE):
            break

        result_lines.append(line)

    return " ".join(result_lines).strip()


def extract_from_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]

    # Render untuk scale XY
    pix = page.get_pixmap(dpi=200)
    pdf_w, pdf_h = page.rect.width, page.rect.height
    png_w, png_h = pix.width, pix.height

    # ==========================
    #  EKSTRAKSI KOORDINAT (XY)
    # ==========================
    nomor        = extract_by_xy(page, 115, 307, 432, 350, pdf_w, pdf_h, png_w, png_h)
    jenis_pph    = extract_by_xy(page, 270, 670, 1580, 715, pdf_w, pdf_h, png_w, png_h)
    dpp          = extract_by_xy(page, 877, 840, 1076, 890, pdf_w, pdf_h, png_w, png_h)
    tarif        = extract_by_xy(page, 1075, 840, 1272, 890, pdf_w, pdf_h, png_w, png_h)
    pph          = extract_by_xy(page, 1270, 840, 1580, 890, pdf_w, pdf_h, png_w, png_h)

    # ==========================
    # FULL TEXT UNTUK REGEX
    # ==========================
    text_all = page.get_text("text")

    # ==========================
    #  OVERRIDE DENGAN REGEX JIKA KOSONG
    # ==========================
    nomor_dokumen = find_nomor_dokumen(text_all)
    nama_pemungut = find_nama_pemungut(text_all)

    # Jika XY kosong → tetap isi dari regex
    if not nomor:
        nomor = nomor_dokumen

    # Return sesuai yang kamu panggil (7 item)
    return nomor, jenis_pph, dpp, tarif, pph, nomor_dokumen, nama_pemungut


# ===================================================
#  PROSES EKSTRAKSI + PROGRESS BAR
# ===================================================
if uploaded_files:
    data_list = []

    st.info("Uploading & Processing files...")

    total = len(uploaded_files)
    progress = st.progress(0)
    status_text = st.empty()

    for i, file in enumerate(uploaded_files):
        status_text.write(f"Processing: **{i+1}/{total}**")

        nomor, jenis_pph, dpp, tarif, pph, nomor_dokumen, nama_pemungut = extract_from_pdf(file.read())

        data_list.append({
            "Nama File": file.name,
            "ID": nomor,
            "Nama Pemungut": nama_pemungut,
            "DPP": dpp,
            "PPH": pph,
            "Tarif": tarif,
            "Jenis PPH": jenis_pph,
            "Nomor Dokumen": nomor_dokumen
        })

        progress.progress((i + 1) / total)

    
    # Streamlit version
    jumlah_file = len(uploaded_files)
    status_text.write(f"✅ Successfully processed {jumlah_file} files")
    progress.progress(1.0)

    df = pd.DataFrame(data_list)

    # ==========================
    # INDEX MULAI DARI 1
    # ==========================
    df.index = df.index + 1

    st.markdown('<div class="small-subheader">Extraction Results</div>', unsafe_allow_html=True)
    st.dataframe(df, use_container_width=True)

    # DOWNLOAD EXCEL
    def to_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='openpyxl')
        df.to_excel(writer, index=False, sheet_name='Data')
        writer.close()
        return output.getvalue()

    excel_file = to_excel(df)

    st.download_button(
        label="Download Extracted File",
        data=excel_file,
        file_name="Extracted File.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # ===================================================
    #  TAMPILKAN LABEL RENAME (SETELAH DOWNLOAD)
    # ===================================================

    st.markdown("""
    <div style='text-align:left; padding: 10px 0;'>
        <p style='font-size:16px; color:#555;'>
            Rename PDF Files Based on Extracted ID
        </p>
        <hr style='margin-top:15px;'>
    </div>
    """, unsafe_allow_html=True)
    # ============================================================
    #  SIAPKAN PDF + BUAT NAMA BARU BERDASARKAN ID
    # ============================================================
    renamed_files = []

    for i, item in enumerate(data_list):
        original_pdf = uploaded_files[i]

        original_pdf.seek(0)
        pdf_bytes = original_pdf.read()

        new_name = f"{item['ID']}.pdf" if item["ID"] else item["Nama File"]

        renamed_files.append({
            "New_Name": new_name,
            "PDF_Bytes": pdf_bytes
        })

    # ============================================================
    #  BUAT ZIP berisi semua PDF hasil rename
    # ============================================================
    import zipfile
    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for file in renamed_files:
            z.writestr(file["New_Name"], file["PDF_Bytes"])

    zip_buffer.seek(0)
    # ============================================================
    #  TOMBOL DOWNLOAD ZIP
    # ============================================================
    st.download_button(
        label="Download All Renamed PDFs (ZIP)",
        data=zip_buffer,
        file_name="Renamed_PDFs.zip",
        mime="application/zip"
    )
try:
    import fitz  # Lokal
except ImportError:
    import pymupdf as fitz  # Streamlit Cloud

import streamlit as st
import pandas as pd
from io import BytesIO


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
    <p style='font-size:16px; color:#555;'>This tool uses Artificial Intelligence Powered Natural Language Processing
        to accurately extract tax data from PDF documents using predefined
        coordinate-based regions</p>
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


def extract_from_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]

    pix = page.get_pixmap(dpi=200)
    pdf_w, pdf_h = page.rect.width, page.rect.height
    png_w, png_h = pix.width, pix.height

    nomor = extract_by_xy(page, 150, 320, 550, 360, pdf_w, pdf_h, png_w, png_h)
    jenis_pph = extract_by_xy(page, 270, 690, 500, 730, pdf_w, pdf_h, png_w, png_h)
    dpp = extract_by_xy(page, 900, 850, 1100, 900, pdf_w, pdf_h, png_w, png_h)
    tarif = extract_by_xy(page, 1100, 850, 1300, 900, pdf_w, pdf_h, png_w, png_h)
    pph = extract_by_xy(page, 1300, 850, 1700, 900, pdf_w, pdf_h, png_w, png_h)
    nomor_dokumen = extract_by_xy(page, 750, 1100, 1600, 1130, pdf_w, pdf_h, png_w, png_h)
    nama_pemungut = extract_by_xy(page, 630, 1360, 1600, 1430, pdf_w, pdf_h, png_w, png_h)

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
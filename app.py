import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import datetime
import zipfile

# =============================================================================
# 1. KONFIGURASI HALAMAN
# =============================================================================
st.set_page_config(page_title="Sistem Diklat DJBC Online", layout="wide", page_icon="📊")

# --- CSS CUSTOM ---
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            .viewerBadge_container__1QSob {display: none !important;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# =============================================================================
# 2. FUNGSI LOGIKA (BACKEND)
# =============================================================================

def set_repeat_table_header(row):
    tr = row._tr; trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader'); tblHeader.set(qn('w:val'), "true")
    trPr.append(tblHeader)

# Fungsi membuat 1 Dokumen Word (Modular)
def create_single_document(row, judul, tgl_pel, tempat_pel, nama_ttd, jabatan_ttd):
    JENIS_FONT = 'Arial'; UKURAN_FONT = 11
    doc = Document()
    style = doc.styles['Normal']; style.font.name = JENIS_FONT; style.font.size = Pt(UKURAN_FONT)
    
    section = doc.sections[0]
    section.top_margin = Cm(2.54); section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54); section.right_margin = Cm(2.54)

    # --- KOP ---
    header_table = doc.add_table(rows=4, cols=3); header_table.autofit = False; header_table.alignment = WD_TABLE_ALIGNMENT.RIGHT 
    header_table.columns[0].width = Cm(1.5); header_table.columns[1].width = Cm(0.3); header_table.columns[2].width = Cm(4.5)
    def isi_sel(r, c, text, size=9, bold=False):
        cell = header_table.cell(r, c); p = cell.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
        run = p.add_run(text); run.font.name = JENIS_FONT; run.font.size = Pt(size); run.bold = bold
        return cell
    c = isi_sel(0, 0, "LAMPIRAN II", 11); c.merge(header_table.cell(0, 2))
    c = isi_sel(1, 0, f"Nota Dinas {jabatan_ttd}", 9); c.merge(header_table.cell(1, 2))
    
    no_nd = row.get('NOMOR_ND', '-') if 'NOMOR_ND' in row else '...................'
    tgl_nd = row.get('TANGGAL_ND', '-') if 'TANGGAL_ND' in row else '...................'
    
    isi_sel(2, 0, "Nomor"); isi_sel(2, 1, ":"); isi_sel(2, 2, str(no_nd))
    isi_sel(3, 0, "Tanggal"); isi_sel(3, 1, ":"); isi_sel(3, 2, str(tgl_nd))

    doc.add_paragraph(""); p = doc.add_paragraph("DAFTAR PESERTA PELATIHAN"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]; run.bold = True; run.font.name = JENIS_FONT; run.font.size = Pt(12) 
    
    # --- INFO ---
    info_table = doc.add_table(rows=3, cols=3); info_table.autofit = False
    info_table.columns[0].width = Cm(4.0); info_table.columns[1].width = Cm(0.5); info_table.columns[2].width = Cm(11.5)
    infos = [("Nama Pelatihan", judul), ("Tanggal", tgl_pel), ("Penyelenggara", tempat_pel)]
    for r, (l, v) in enumerate(infos): info_table.cell(r,0).text = l; info_table.cell(r,1).text = ":"; info_table.cell(r,2).text = v
    for r_obj in info_table.rows:
        for cell in r_obj.cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_after = Pt(2)
                if p.runs: p.runs[0].font.name = JENIS_FONT; p.runs[0].font.size = Pt(11)
    doc.add_paragraph("")

    # --- TABEL DATA ---
    table = doc.add_table(rows=2, cols=5); table.style = 'Table Grid'; table.autofit = False
    headers = ['NO', 'NAMA PEGAWAI', 'NIP', 'PANGKAT - GOL', 'SATUAN KERJA']
    widths = [Cm(1.0), Cm(5.0), Cm(3.8), Cm(3.5), Cm(3.5)]
    hdr = table.rows[0].cells
    for i in range(5):
        hdr[i].text = headers[i]; hdr[i].width = widths[i]
        p = hdr[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]; run.bold = True; run.font.name = JENIS_FONT; run.font.size = Pt(10)

    # Isi Data
    vals = [row['NO'], row['NAMA'], row['NIP'], row['PANGKAT'], row['SATKER']]
    row_cells = table.rows[1].cells
    for i in range(5):
        row_cells[i].width = widths[i]; row_cells[i].text = str(vals[i])
        row_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = row_cells[i].paragraphs[0]; p.paragraph_format.space_after = Pt(2)
        run = p.runs[0]; run.font.name = JENIS_FONT; run.font.size = Pt(10)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in [0, 2] else WD_ALIGN_PARAGRAPH.LEFT

    # --- TTD ---
    doc.add_paragraph(""); ttd_table = doc.add_table(rows=1, cols=2); ttd_table.autofit = False
    ttd_table.columns[0].width = Cm(9.0); ttd_table.columns[1].width = Cm(7.0)
    cell = ttd_table.cell(0, 1); p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.line_spacing = 1.0 
    run = p.add_run(f"{jabatan_ttd},"); run.font.name = JENIS_FONT; run.font.size = Pt(11); p.add_run("\n" * 6)
    run = p.add_run("Ditandatangani secara elektronik"); run.font.name = JENIS_FONT; run.font.size = Pt(9); run.font.color.rgb = RGBColor(150, 150, 150); p.add_run("\n")
    run = p.add_run(nama_ttd); run.font.name = JENIS_FONT; run.font.size = Pt(11); run.bold = False 

    # Save to buffer
    f_out = io.BytesIO()
    doc.save(f_out)
    f_out.seek(0)
    return f_out

def generate_zip_files(df, nama_ttd, jabatan_ttd):
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for idx, row in df.iterrows():
            judul = row.get('JUDUL_PELATIHAN', 'Diklat')
            tgl = row.get('TANGGAL_PELATIHAN', '-')
            tempat = row.get('TEMPAT', '-')
            clean_nama = str(row['NAMA']).replace('/', '_').replace('\\', '_')
            nama_file = f"{clean_nama}_{str(row['NIP'])}.docx"
            
            doc_buffer = create_single_document(row, judul, tgl, tempat, nama_ttd, jabatan_ttd)
            zip_file.writestr(nama_file, doc_buffer.getvalue())
            
    zip_buffer.seek(0)
    return zip_buffer

def generate_word_combined(df, nama_ttd, jabatan_ttd):
    output = io.BytesIO()
    JENIS_FONT = 'Arial'; UKURAN_FONT = 11
    doc = Document()
    style = doc.styles['Normal']; style.font.name = JENIS_FONT; style.font.size = Pt(UKURAN_FONT)
    section = doc.sections[0]; section.top_margin = Cm(2.54); section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54); section.right_margin = Cm(2.54)
    
    kelompok = df.groupby('JUDUL_PELATIHAN')
    counter = 0; total = len(kelompok)
    for judul, group in kelompok:
        for i_peserta, row_peserta in group.iterrows():
            counter += 1
            header_table = doc.add_table(rows=4, cols=3); header_table.alignment = WD_TABLE_ALIGNMENT.RIGHT 
            header_table.columns[0].width = Cm(1.5); header_table.columns[2].width = Cm(4.5)
            
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(f"LAMPIRAN II\nNota Dinas {jabatan_ttd}")
            r.font.size = Pt(9); r.font.name = JENIS_FONT
            
            doc.add_paragraph("")
            p = doc.add_paragraph("DAFTAR PESERTA PELATIHAN"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.runs[0].bold=True; p.runs[0].font.size=Pt(12); p.runs[0].font.name=JENIS_FONT
            
            p_info = doc.add_paragraph()
            p_info.add_run(f"Nama Pelatihan : {judul}\n").font.name = JENIS_FONT
            p_info.add_run(f"Tanggal        : {row_peserta['TANGGAL_PELATIHAN']}\n").font.name = JENIS_FONT
            p_info.add_run(f"Tempat         : {row_peserta['TEMPAT']}").font.name = JENIS_FONT
            
            doc.add_paragraph("")
            tbl = doc.add_table(rows=2, cols=5); tbl.style='Table Grid'
            hdrs = ['NO', 'NAMA', 'NIP', 'PANGKAT', 'SATKER']
            for i, h in enumerate(hdrs): 
                r = tbl.rows[0].cells[i].paragraphs[0].add_run(h); r.bold=True; r.font.size=Pt(10); r.font.name=JENIS_FONT
                tbl.rows[0].cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            vals = [row_peserta['NO'], row_peserta['NAMA'], row_peserta['NIP'], row_peserta['PANGKAT'], row_peserta['SATKER']]
            for i, v in enumerate(vals):
                r = tbl.rows[1].cells[i].paragraphs[0].add_run(str(v)); r.font.size=Pt(10); r.font.name=JENIS_FONT

            doc.add_paragraph("")
            p = doc.add_paragraph(f"\n\nDitandatangani secara elektronik\n{nama_ttd}"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.runs[0].font.name = JENIS_FONT
            
            if counter < len(df): doc.add_page_break()

    doc.save(output)
    output.seek(0)
    return output


# =============================================================================
# 3. GUI UTAMA
# =============================================================================

st.title("Sistem Administrasi Diklat DJBC 🇮🇩")
st.markdown("---")

# SIDEBAR
with st.sidebar:
    st.header("📂 Upload Data")
    uploaded_file = st.file_uploader("Upload Excel", type=['xlsx'])
    
    st.markdown("### ✍️ Penanda Tangan")
    nama_ttd = st.text_input("Nama Pejabat", "Ayu Sukorini")
    jabatan_ttd = st.text_input("Jabatan", "Sekretaris Direktorat Jenderal")

    st.markdown("---")
    if st.button("📥 Download Template"):
        df_dummy = pd.DataFrame({
            "JUDUL_PELATIHAN": ["Diklat Teknis A", "Diklat Teknis A"],
            "TANGGAL_PELATIHAN": ["10-12 Jan 2025", "10-12 Jan 2025"],
            "TEMPAT": ["Pusdiklat BC", "Pusdiklat BC"],
            "NO": [1, 2],
            "NAMA PEGAWAI": ["Fajar Ali", "Dede Kurnia"],
            "NIP": ["19990101...", "19950505..."],
            "PANGKAT - GOL": ["Pengatur (II/c)", "Penata Muda (III/a)"],
            "SATUAN KERJA": ["KPU Batam", "KPPBC Jakarta"],
            "NOMOR_ND": ["ND-123/2025", "ND-123/2025"],
            "TANGGAL_ND": ["10 Januari 2025", "10 Januari 2025"]
        })
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_dummy.to_excel(writer, index=False)
        buffer.seek(0)
        st.download_button("Klik untuk Download", buffer, "Template_Peserta.xlsx")

# MAIN AREA
if uploaded_file:
    try:
        # BACA DATA
        df_raw = pd.read_excel(uploaded_file, dtype=str)
        # RENAMING AGAR KONSISTEN
        df_raw = df_raw.rename(columns={
            'NAMA PEGAWAI': 'NAMA', 
            'PANGKAT - GOL': 'PANGKAT', 
            'PANGKAT GOL': 'PANGKAT', # Jaga-jaga variasi nama kolom
            'SATUAN KERJA': 'SATKER'
        })
        # Standarisasi kolom jadi huruf besar & tanpa spasi
        df_raw.columns = [c.strip().upper().replace(" ", "_") for c in df_raw.columns]
        df_raw = df_raw.fillna("-")
        
        tab1, tab2 = st.tabs(["📝 Generator Dokumen", "📊 Dashboard"])
        
        with tab1:
            st.success(f"Terbaca: **{len(df_raw)} Data**.")
            st.info("💡 Edit data di bawah jika ada typo, lalu pilih jenis download.")
            
            df_edited = st.data_editor(df_raw, num_rows="dynamic", use_container_width=True)
            
            st.markdown("### ⚡ Pilihan Output")
            col1, col2 = st.columns(2)
            ts = datetime.datetime.now().strftime("%H%M%S")
            
            with col1:
                st.markdown("#### 📄 Satu File Gabungan")
                st.caption("Semua peserta dalam 1 file Word panjang.")
                if st.button("Generate Single Word"):
                    with st.spinner("Membuat dokumen..."):
                        docx_file = generate_word_combined(df_edited, nama_ttd, jabatan_ttd)
                        st.download_button("📥 Download .docx", docx_file, f"Lampiran_Gabungan_{ts}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            with col2:
                st.markdown("#### 📦 File Terpisah (.ZIP)")
                st.caption("Setiap peserta punya file Word sendiri-sendiri.")
                if st.button("Generate Bulk ZIP", type="primary"):
                    with st.spinner("Memisahkan file per peserta..."):
                        zip_file = generate_zip_files(df_edited, nama_ttd, jabatan_ttd)
                        st.download_button("📥 Download .zip", zip_file, f"Arsip_Peserta_{ts}.zip", "application/zip")

        with tab2:
            df_viz = df_edited
            
            # 1. METRIK UTAMA
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Peserta", len(df_viz))
            c2.metric("Total Satker", df_viz['SATKER'].nunique())
            c3.metric("Total Diklat", df_viz['JUDUL_PELATIHAN'].nunique())
            
            st.markdown("---")
            
            # 2. GRAFIK BARIS 1 (Satker & Pangkat)
            c_g1, c_g2 = st.columns(2)
            
            with c_g1:
                st.subheader("🏢 Top 10 Satuan Kerja")
                try:
                    top_satker = df_viz['SATKER'].value_counts().head(10).sort_values()
                    fig, ax = plt.subplots(figsize=(5,4))
                    top_satker.plot(kind='barh', ax=ax, color='#3498db')
                    ax.set_ylabel("")
                    st.pyplot(fig)
                    plt.close(fig) # Hemat memori
                except Exception as e: st.warning(f"Gagal memuat grafik Satker: {e}")
            
            with c_g2:
                st.subheader("👮 Komposisi Pangkat")
                try:
                    # Cek kolom Pangkat
                    if 'PANGKAT' in df_viz.columns:
                        pangkat_col = 'PANGKAT'
                    elif 'PANGKAT_GOL' in df_viz.columns: # Fallback
                        pangkat_col = 'PANGKAT_GOL'
                    else:
                        pangkat_col = None
                    
                    if pangkat_col:
                        pangkat_counts = df_viz[pangkat_col].value_counts()
                        if not pangkat_counts.empty:
                            fig2, ax2 = plt.subplots(figsize=(5,4))
                            ax2.pie(pangkat_counts, labels=pangkat_counts.index, autopct='%1.1f%%', startangle=90)
                            st.pyplot(fig2)
                            plt.close(fig2)
                        else: st.warning("Data Pangkat kosong.")
                    else: st.warning("Kolom 'Pangkat' tidak ditemukan.")
                except Exception as e: st.warning(f"Gagal memuat Pie Chart: {e}")

            st.markdown("---")

            # 3. GRAFIK BARIS 2 (Lokasi & Tanggal - BARU)
            c_g3, c_g4 = st.columns(2)

            with c_g3:
                st.subheader("📍 Lokasi Pelaksanaan")
                try:
                    if 'TEMPAT' in

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

# --- 1. KONFIGURASI HALAMAN (WAJIB PALING ATAS) ---
st.set_page_config(page_title="Sistem Diklat DJBC Online", layout="wide", page_icon="📝")

# --- CSS CUSTOM ---
# Menyembunyikan elemen-elemen bawaan Streamlit agar terlihat bersih
hide_st_style = """
            <style>
            /* 1. Menghilangkan Menu Hamburger (Kanan Atas) */
            #MainMenu {visibility: hidden;}
            
            /* 2. Menghilangkan Header Atas */
            header {visibility: hidden;}
            
            /* 3. Menghilangkan Footer Standar */
            footer {visibility: hidden;}
            
            /* 4. MENGHILANGKAN LOGO STREAMLIT (Kanan Bawah) */
            /* Target spesifik untuk Viewer Badge */
            .viewerBadge_container__1QSob {display: none !important;}
            .viewerBadge_link__1S137 {display: none !important;}
            
            /* Target cadangan jika nama class berubah */
            [data-testid="stStatusWidget"] {visibility: hidden !important;}
            
            /* Menghilangkan whitespace di bawah yang kadang muncul */
            div.block-container {padding-bottom: 1rem;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# =============================================================================
# FUNGSI LOGIKA (BACKEND)
# =============================================================================

def set_repeat_table_header(row):
    """Agar header tabel berulang jika pindah halaman"""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader')
    tblHeader.set(qn('w:val'), "true")
    trPr.append(tblHeader)

def add_page_number(doc):
    """Menambahkan nomor halaman di footer"""
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar1)
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    run._r.append(instrText)
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar2)

def generate_word_in_memory(df, nama_ttd, jabatan_ttd):
    """Fungsi Utama Generator Word (Disimpan ke Memory RAM)"""
    JENIS_FONT = 'Arial'
    UKURAN_FONT = 11
    
    # Buffer in-memory (pengganti file fisik)
    output = io.BytesIO()
    
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = JENIS_FONT
    style.font.size = Pt(UKURAN_FONT)
    
    # Margin
    section = doc.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)
    
    add_page_number(doc)

    # Grouping Data
    kelompok_data = df.groupby('JUDUL_PELATIHAN')
    counter = 0
    total_grup = len(kelompok_data)
    
    for judul, data_grup in kelompok_data:
        counter += 1
        tgl_pel = data_grup.iloc[0]['TANGGAL_PELATIHAN']
        tempat_pel = data_grup.iloc[0]['TEMPAT']

        # --- TABEL KOP SURAT (LAMPIRAN) ---
        header_table = doc.add_table(rows=4, cols=3)
        header_table.autofit = False
        header_table.alignment = WD_TABLE_ALIGNMENT.RIGHT 
        header_table.columns[0].width = Cm(1.5)
        header_table.columns[1].width = Cm(0.3)
        header_table.columns[2].width = Cm(4.5)
        
        def isi_sel(r, c, text, size=9, bold=False):
            cell = header_table.cell(r, c)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(text)
            run.font.name = JENIS_FONT
            run.font.size = Pt(size)
            run.bold = bold
            return cell

        c = isi_sel(0, 0, "LAMPIRAN II", 11)
        c.merge(header_table.cell(0, 2))
        
        c = isi_sel(1, 0, f"Nota Dinas {jabatan_ttd}", 9)
        c.merge(header_table.cell(1, 2))
        
        data_h = [("Nomor", ":", "[@NomorND]"), ("Tanggal", ":", "[@TanggalND]")]
        for i, (l, s, v) in enumerate(data_h):
            isi_sel(i+2, 0, l)
            isi_sel(i+2, 1, s)
            isi_sel(i+2, 2, v)

        doc.add_paragraph("")
        p = doc.add_paragraph("DAFTAR PESERTA PELATIHAN")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]
        run.bold = True
        run.font.name = JENIS_FONT
        run.font.size = Pt(12) 
        
        # --- TABEL INFO DIKLAT ---
        info_table = doc.add_table(rows=3, cols=3)
        info_table.autofit = False
        info_table.columns[0].width = Cm(4.0)
        info_table.columns[1].width = Cm(0.5)
        info_table.columns[2].width = Cm(11.5)
        
        infos = [("Nama Pelatihan", judul), ("Tanggal", tgl_pel), ("Penyelenggara", tempat_pel)]
        for r, (l, v) in enumerate(infos):
            info_table.cell(r,0).text = l
            info_table.cell(r,1).text = ":"
            info_table.cell(r,2).text = v
        
        for row in info_table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.space_after = Pt(2)
                    if p.runs:
                        p.runs[0].font.name = JENIS_FONT
                        p.runs[0].font.size = Pt(11)
        doc.add_paragraph("")

        # --- TABEL DATA PESERTA ---
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        table.autofit = False
        
        headers = ['NO', 'NAMA PEGAWAI', 'NIP', 'PANGKAT - GOL', 'SATUAN KERJA']
        widths = [Cm(1.0), Cm(5.0), Cm(3.8), Cm(3.5), Cm(3.5)]
        
        hdr = table.rows[0].cells
        set_repeat_table_header(table.rows[0])
        
        for i in range(5):
            hdr[i].text = headers[i]
            hdr[i].width = widths[i]
            p = hdr[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.runs[0]
            run.bold = True
            run.font.name = JENIS_FONT
            run.font.size = Pt(10)

        for idx, row in data_grup.iterrows():
            rc = table.add_row().cells
            vals = [row['NO'], row['NAMA'], row['NIP'], row['PANGKAT'], row['SATKER']]
            for i in range(5):
                rc[i].width = widths[i]
                rc[i].text = str(vals[i])
                rc[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = rc[i].paragraphs[0]
                p.paragraph_format.space_after = Pt(2)
                run = p.runs[0]
                run.font.name = JENIS_FONT
                run.font.size = Pt(10)
                if i in [0, 2]: # No & NIP Center
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        if counter < total_grup:
            doc.add_page_break()

    # --- TANDA TANGAN ---
    doc.add_paragraph("")
    ttd_table = doc.add_table(rows=1, cols=2)
    ttd_table.autofit = False
    ttd_table.columns[0].width = Cm(9.0)
    ttd_table.columns[1].width = Cm(7.0)
    
    cell = ttd_table.cell(0, 1)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.0 
    
    run = p.add_run(f"{jabatan_ttd},")
    run.font.name = JENIS_FONT
    run.font.size = Pt(11)
    p.add_run("\n" * 8)
    run = p.add_run("Ditandatangani secara elektronik")
    run.font.name = JENIS_FONT
    run.font.size = Pt(9)
    run.italic = False
    run.font.color.rgb = RGBColor(150, 150, 150)
    p.add_run("\n")
    run = p.add_run(nama_ttd)
    run.font.name = JENIS_FONT
    run.font.size = Pt(11)
    run.bold = False 

    # Simpan ke memory buffer
    doc.save(output)
    output.seek(0)
    return output

# =============================================================================
# GUI STREAMLIT (FRONTEND)
# =============================================================================

st.title("Sistem Administrasi Diklat DJBC 🇮🇩")
st.markdown("---")

# --- SIDEBAR (UPLOAD & INPUT) ---
with st.sidebar:
    st.header("📂 Upload Data")
    uploaded_file = st.file_uploader("Upload File Excel Peserta", type=['xlsx'])
    
    st.markdown("### ✍️ Penanda Tangan")
    nama_ttd = st.text_input("Nama Pejabat", "Ayu Sukorini")
    jabatan_ttd = st.text_input("Jabatan", "Sekretaris Direktorat Jenderal")

    st.markdown("---")
    if st.button("📥 Download Template Excel"):
        # Membuat dummy file di memory
        df_dummy = pd.DataFrame({
            "JUDUL_PELATIHAN": ["Diklat Teknis A", "Diklat Teknis A"],
            "TANGGAL_PELATIHAN": ["10-12 Jan 2025", "10-12 Jan 2025"],
            "TEMPAT": ["Pusdiklat BC", "Pusdiklat BC"],
            "NO": [1, 2],
            "NAMA PEGAWAI": ["Fajar Ali", "Dede Kurnia"],
            "NIP": ["19990101...", "19950505..."],
            "PANGKAT - GOL": ["Pengatur (II/c)", "Penata Muda (III/a)"],
            "SATUAN KERJA": ["KPU Batam", "KPPBC Jakarta"],
            "EMAIL": ["fajar@kemenkeu.go.id", "dede@customs.go.id"]
        })
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_dummy.to_excel(writer, index=False)
        buffer.seek(0)
        st.download_button(
            label="Klik untuk Download Template",
            data=buffer,
            file_name="Template_Peserta.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- AREA UTAMA ---
if uploaded_file:
    try:
        # Baca Excel
        df_raw = pd.read_excel(uploaded_file, dtype=str)
        # Normalisasi Nama Kolom
        df_raw = df_raw.rename(columns={'NAMA PEGAWAI': 'NAMA', 'PANGKAT - GOL': 'PANGKAT', 'SATUAN KERJA': 'SATKER'})
        df_raw.columns = [c.strip().upper().replace(" ", "_") for c in df_raw.columns]
        df_raw = df_raw.fillna("-")
        
        # TABS UTAMA
        tab1, tab2 = st.tabs(["📝 Generator (Live Editor)", "📊 Dashboard Analitik"])
        
        # --- TAB 1: LIVE EDITOR ---
        with tab1:
            st.success(f"File berhasil dibaca: **{len(df_raw)} Data** terdeteksi.")
            
            st.markdown("### ✏️ Live Editor")
            st.caption("Klik dua kali pada sel tabel di bawah ini untuk mengedit data (Typo nama, NIP, dll).")
            
            # WIDGET DATA EDITOR (FITUR UTAMA)
            df_edited = st.data_editor(df_raw, num_rows="dynamic", use_container_width=True)
            
            st.markdown("---")
            col_btn, col_info = st.columns([1, 2])
            
            with col_btn:
                # Tombol Generate
                if st.button("⚡ GENERATE DOKUMEN WORD", type="primary"):
                    with st.spinner("Sedang memproses dokumen..."):
                        try:
                            # Generate menggunakan data hasil editan (df_edited)
                            docx_file = generate_word_in_memory(df_edited, nama_ttd, jabatan_ttd)
                            
                            ts = datetime.datetime.now().strftime("%H%M%S")
                            filename = f"Lampiran_Peserta_{ts}.docx"
                            
                            st.success("✅ Dokumen Selesai!")
                            st.download_button(
                                label="📥 Download Hasil (.docx)",
                                data=docx_file,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        except Exception as e:
                            st.error(f"Terjadi Kesalahan: {e}")

        # --- TAB 2: DASHBOARD ---
        with tab2:
            # Gunakan data yang sudah diedit untuk visualisasi juga
            df_viz = df_edited
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Peserta", len(df_viz))
            col2.metric("Total Satker", df_viz['SATKER'].nunique())
            col3.metric("Total Diklat", df_viz['JUDUL_PELATIHAN'].nunique())
            
            st.markdown("---")
            
            c1, c2 = st.columns(2)
            
            with c1:
                st.subheader("Top Satuan Kerja")
                try:
                    top_satker = df_viz['SATKER'].value_counts().head(10).sort_values()
                    fig, ax = plt.subplots()
                    top_satker.plot(kind='barh', ax=ax, color='#3498db')
                    st.pyplot(fig)
                except:
                    st.warning("Data Satker tidak cukup untuk grafik.")
            
            with c2:
                st.subheader("Komposisi Pangkat")
                try:
                    pangkat_counts = df_viz['PANGKAT'].value_counts()
                    fig2, ax2 = plt.subplots()
                    ax2.pie(pangkat_counts, labels=pangkat_counts.index, autopct='%1.1f%%', startangle=90)
                    st.pyplot(fig2)
                except:
                    st.warning("Data Pangkat tidak cukup untuk grafik.")

    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
        st.warning("Pastikan file Excel memiliki format yang sesuai atau install library 'openpyxl' di requirements.txt")

else:
    st.info("👈 Silakan upload file Excel pada menu di sebelah kiri.")


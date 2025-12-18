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

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Sistem Diklat DJBC Online", layout="wide", page_icon="📝")

# --- CSS CUSTOM UNTUK TAMPILAN ---
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    h1 { color: #2c3e50; }
    div.stButton > button:first-child { background-color: #2c3e50; color: white; }
    </style>
""", unsafe_allow_html=True)

# =============================================================================
# FUNGSI UTILITAS (LOGIKA SAMA DENGAN VERSI DESKTOP)
# =============================================================================

def set_repeat_table_header(row):
    tr = row._tr; trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader'); tblHeader.set(qn('w:val'), "true")
    trPr.append(tblHeader)

def add_page_number(doc):
    section = doc.sections[0]; footer = section.footer
    p = footer.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin'); run._r.append(fldChar1)
    instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = "PAGE"; run._r.append(instrText)
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end'); run._r.append(fldChar2)

def generate_word_in_memory(df, nama_ttd, jabatan_ttd):
    """Fungsi ini dimodifikasi untuk Web: Tidak menyimpan file ke disk, tapi ke Memory RAM (Buffer)"""
    JENIS_FONT = 'Arial'; UKURAN_FONT = 11
    
    # In-memory byte buffer
    output = io.BytesIO()
    
    doc = Document()
    style = doc.styles['Normal']; style.font.name = JENIS_FONT; style.font.size = Pt(UKURAN_FONT)
    section = doc.sections[0]
    section.top_margin = Cm(2.54); section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54); section.right_margin = Cm(2.54)
    add_page_number(doc)

    kelompok_data = df.groupby('JUDUL_PELATIHAN')
    counter = 0; total_grup = len(kelompok_data)
    
    for judul, data_grup in kelompok_data:
        counter += 1
        tgl_pel = data_grup.iloc[0]['TANGGAL_PELATIHAN']
        tempat_pel = data_grup.iloc[0]['TEMPAT']

        # --- HEADER TABLE ---
        header_table = doc.add_table(rows=4, cols=3); header_table.autofit = False; header_table.alignment = WD_TABLE_ALIGNMENT.RIGHT 
        header_table.columns[0].width = Cm(1.5); header_table.columns[1].width = Cm(0.3); header_table.columns[2].width = Cm(4.5)
        
        def isi_sel(r, c, text, size=9, bold=False):
            cell = header_table.cell(r, c); p = cell.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
            run = p.add_run(text); run.font.name = JENIS_FONT; run.font.size = Pt(size); run.bold = bold
            return cell

        c = isi_sel(0, 0, "LAMPIRAN II", 11); c.merge(header_table.cell(0, 2))
        c = isi_sel(1, 0, f"Nota Dinas {jabatan_ttd}", 9); c.merge(header_table.cell(1, 2))
        data_h = [("Nomor", ":", "[@NomorND]"), ("Tanggal", ":", "[@TanggalND]")]
        for i, (l, s, v) in enumerate(data_h): isi_sel(i+2, 0, l); isi_sel(i+2, 1, s); isi_sel(i+2, 2, v)

        doc.add_paragraph("")
        p = doc.add_paragraph("DAFTAR PESERTA PELATIHAN"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]; run.bold = True; run.font.name = JENIS_FONT; run.font.size = Pt(12) 
        
        info_table = doc.add_table(rows=3, cols=3); info_table.autofit = False
        info_table.columns[0].width = Cm(4.0); info_table.columns[1].width = Cm(0.5); info_table.columns[2].width = Cm(11.5)
        infos = [("Nama Pelatihan", judul), ("Tanggal", tgl_pel), ("Penyelenggara", tempat_pel)]
        for r, (l, v) in enumerate(infos): info_table.cell(r,0).text = l; info_table.cell(r,1).text = ":"; info_table.cell(r,2).text = v
        for row in info_table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.space_after = Pt(2); 
                    if p.runs: p.runs[0].font.name = JENIS_FONT; p.runs[0].font.size = Pt(11)
        doc.add_paragraph("")

        # --- DATA TABLE ---
        table = doc.add_table(rows=1, cols=5); table.style = 'Table Grid'; table.autofit = False
        headers = ['NO', 'NAMA PEGAWAI', 'NIP', 'PANGKAT - GOL', 'SATUAN KERJA']
        widths = [Cm(1.0), Cm(5.0), Cm(3.8), Cm(3.5), Cm(3.5)]
        hdr = table.rows[0].cells; set_repeat_table_header(table.rows[0])
        for i in range(5):
            hdr[i].text = headers[i]; hdr[i].width = widths[i]
            p = hdr[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.runs[0]; run.bold = True; run.font.name = JENIS_FONT; run.font.size = Pt(10)

        for idx, row in data_grup.iterrows():
            rc = table.add_row().cells; vals = [row['NO'], row['NAMA'], row['NIP'], row['PANGKAT'], row['SATKER']]
            for i in range(5):
                rc[i].width = widths[i]; rc[i].text = str(vals[i]); rc[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = rc[i].paragraphs[0]; p.paragraph_format.space_after = Pt(2)
                run = p.runs[0]; run.font.name = JENIS_FONT; run.font.size = Pt(10)
                if i in [0, 2]: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else: p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        if counter < total_grup: doc.add_page_break()

    # --- TTD ---
    doc.add_paragraph(""); ttd_table = doc.add_table(rows=1, cols=2); ttd_table.autofit = False
    ttd_table.columns[0].width = Cm(9.0); ttd_table.columns[1].width = Cm(7.0)
    cell = ttd_table.cell(0, 1); p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.line_spacing = 1.0 
    
    run = p.add_run(f"{jabatan_ttd},"); run.font.name = JENIS_FONT; run.font.size = Pt(11)
    p.add_run("\n" * 8)
    run = p.add_run("Ditandatangani secara elektronik"); run.font.name = JENIS_FONT; run.font.size = Pt(9); run.italic = False; run.font.color.rgb = RGBColor(150, 150, 150)
    p.add_run("\n")
    run = p.add_run(nama_ttd); run.font.name = JENIS_FONT; run.font.size = Pt(11); run.bold = False 

    doc.save(output)
    output.seek(0)
    return output

# =============================================================================
# GUI STREAMLIT
# =============================================================================

st.title("Sistem Administrasi Diklat DJBC 🇮🇩")
st.markdown("---")

# SIDEBAR (UPLOAD FILE)
with st.sidebar:
    st.header("📂 Upload Data")
    uploaded_file = st.file_uploader("Upload File Excel Peserta", type=['xlsx'])
    
    st.markdown("### ✍️ Penanda Tangan")
    nama_ttd = st.text_input("Nama Pejabat", "Ayu Sukorini")
    jabatan_ttd = st.text_input("Jabatan", "Sekretaris Direktorat Jenderal")

    if st.button("📥 Download Template Excel"):
        # Create dummy template in memory
        df_dummy = pd.DataFrame({
            "JUDUL_PELATIHAN": ["Diklat Teknis A"], "TANGGAL_PELATIHAN": ["10-12 Jan 2025"],
            "TEMPAT": ["Pusdiklat BC"], "NO": [1], "NAMA PEGAWAI": ["Contoh Nama"],
            "NIP": ["19999..."], "PANGKAT - GOL": ["II/c"], "SATUAN KERJA": ["KPU Batam"],
            "EMAIL": ["contoh@kemenkeu.go.id"]
        })
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_dummy.to_excel(writer, index=False)
        buffer.seek(0)
        st.download_button("Klik untuk Download", buffer, "Template_Peserta.xlsx")

# MAIN CONTENT
if uploaded_file:
    # BACA DATA
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
        df = df.rename(columns={'NAMA PEGAWAI': 'NAMA', 'PANGKAT - GOL': 'PANGKAT', 'SATUAN KERJA': 'SATKER'})
        df.columns = [c.strip().upper().replace(" ", "_") for c in df.columns]
        df = df.fillna("-")
        
        # TABS
        tab1, tab2 = st.tabs(["📝 Generator Dokumen", "📊 Dashboard Analitik"])
        
        # --- TAB 1: GENERATOR ---
        with tab1:
            st.info(f"File berhasil dibaca: **{len(df)} Peserta** terdeteksi.")
            st.dataframe(df.head())
            
            if st.button("⚡ GENERATE LAMPIRAN WORD", type="primary"):
                with st.spinner("Sedang membuat dokumen..."):
                    try:
                        # Panggil fungsi generate (memory)
                        docx_file = generate_word_in_memory(df, nama_ttd, jabatan_ttd)
                        
                        ts = datetime.datetime.now().strftime("%H%M%S")
                        filename = f"Lampiran_Peserta_{ts}.docx"
                        
                        st.success("Dokumen Selesai Dibuat!")
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
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Peserta", len(df))
            col2.metric("Total Satker", df['SATKER'].nunique())
            col3.metric("Total Diklat", df['JUDUL_PELATIHAN'].nunique())
            
            st.markdown("---")
            
            c1, c2 = st.columns(2)
            
            with c1:
                st.subheader("Top Satuan Kerja")
                top_satker = df['SATKER'].value_counts().head(10).sort_values()
                fig, ax = plt.subplots()
                top_satker.plot(kind='barh', ax=ax, color='#3498db')
                st.pyplot(fig)
            
            with c2:
                st.subheader("Komposisi Pangkat")
                pangkat_counts = df['PANGKAT'].value_counts()
                fig2, ax2 = plt.subplots()
                ax2.pie(pangkat_counts, labels=pangkat_counts.index, autopct='%1.1f%%', startangle=90)
                st.pyplot(fig2)

    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")

else:
    st.info("👈 Silakan upload file Excel pada menu di sebelah kiri.")
    st.markdown("""
    ### Panduan Penggunaan:
    1. Upload file Excel Data Peserta.
    2. Isi Nama dan Jabatan Penanda Tangan.
    3. Masuk ke Tab **Generator Dokumen** untuk download Word.
    4. Masuk ke Tab **Dashboard** untuk melihat statistik.
    """)
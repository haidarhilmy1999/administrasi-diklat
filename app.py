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

# --- LIBRARY GOOGLE SHEETS ---
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# =============================================================================
# 1. KONFIGURASI HALAMAN
# =============================================================================
st.set_page_config(page_title="Sistem Diklat DJBC Online", layout="wide", page_icon="📈")

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

# Nama Google Sheet
NAMA_GOOGLE_SHEET = "Database_Diklat_DJBC"

# --- INISIALISASI RIWAYAT LOKAL ---
if 'history_log' not in st.session_state:
    st.session_state['history_log'] = pd.DataFrame(columns=['TIMESTAMP', 'NAMA', 'NIP', 'DIKLAT', 'SATKER'])

# =============================================================================
# 2. FUNGSI GOOGLE SHEETS
# =============================================================================

def connect_to_gsheet():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        sheet = client.open(NAMA_GOOGLE_SHEET).sheet1
        return sheet
    except Exception as e:
        return None

def save_to_cloud_database(df_input):
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        data_to_save = df_input.copy()
        # Normalisasi nama kolom untuk database
        target_cols = ['NAMA', 'NIP', 'JUDUL_PELATIHAN', 'SATKER']
        final_cols = {}
        for target in target_cols:
            for col in data_to_save.columns:
                if target in col: 
                    final_cols[col] = target if target != 'JUDUL_PELATIHAN' else 'DIKLAT'
                    break
        
        if final_cols:
            data_to_save = data_to_save[list(final_cols.keys())].rename(columns=final_cols)
            data_to_save.insert(0, 'TIMESTAMP', current_time)
            
            st.session_state['history_log'] = pd.concat([st.session_state['history_log'], data_to_save], ignore_index=True)
            
            sheet = connect_to_gsheet()
            if sheet:
                data_list = data_to_save.astype(str).values.tolist()
                sheet.append_rows(data_list)
                return "Sukses Cloud"
            else:
                return "Gagal Koneksi Cloud"
        return "Kolom Data Tidak Lengkap"
    except Exception as e:
        return f"Error: {e}"

# =============================================================================
# 3. FUNGSI GENERATOR WORD (PEMBANTU)
# =============================================================================

def set_repeat_table_header(row):
    """Membuat header tabel berulang di tiap halaman"""
    tr = row._tr; trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader'); tblHeader.set(qn('w:val'), "true")
    trPr.append(tblHeader)

# --- FUNGSI 1: GENERATE PER PESERTA (UNTUK ZIP) ---
def create_single_document(row, judul, tgl_pel, tempat_pel, nama_ttd, jabatan_ttd):
    JENIS_FONT = 'Arial'; UKURAN_FONT = 11
    doc = Document()
    style = doc.styles['Normal']; style.font.name = JENIS_FONT; style.font.size = Pt(UKURAN_FONT)
    section = doc.sections[0]; section.top_margin = Cm(2.54); section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54); section.right_margin = Cm(2.54)

    header_table = doc.add_table(rows=4, cols=3); header_table.alignment = WD_TABLE_ALIGNMENT.RIGHT 
    header_table.columns[0].width = Cm(1.5); header_table.columns[2].width = Cm(4.5)
    def isi_sel(r, c, text, size=9, bold=False):
        cell = header_table.cell(r, c); p = cell.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
        run = p.add_run(text); run.font.name = JENIS_FONT; run.font.size = Pt(size); run.bold = bold
    isi_sel(0, 0, "LAMPIRAN II", 11); header_table.cell(0, 2).merge(header_table.cell(0, 0)); header_table.cell(0,0).text="LAMPIRAN II"
    isi_sel(1, 0, f"Nota Dinas {jabatan_ttd}", 9); header_table.cell(1, 2).merge(header_table.cell(1, 0)); header_table.cell(1,0).text=f"Nota Dinas {jabatan_ttd}"
    
    no_nd = row.get('NOMOR_ND', '...................'); tgl_nd = row.get('TANGGAL_ND', '...................')
    isi_sel(2, 0, "Nomor"); isi_sel(2, 1, ":"); isi_sel(2, 2, str(no_nd))
    isi_sel(3, 0, "Tanggal"); isi_sel(3, 1, ":"); isi_sel(3, 2, str(tgl_nd))

    doc.add_paragraph(""); p = doc.add_paragraph("DAFTAR PESERTA PELATIHAN"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]; run.bold = True; run.font.name = JENIS_FONT; run.font.size = Pt(12) 
    
    info_table = doc.add_table(rows=3, cols=3); 
    infos = [("Nama Pelatihan", judul), ("Tanggal", tgl_pel), ("Penyelenggara", tempat_pel)]
    for r, (l, v) in enumerate(infos): info_table.cell(r,0).text = l; info_table.cell(r,1).text = ":"; info_table.cell(r,2).text = v

    doc.add_paragraph("")
    table = doc.add_table(rows=2, cols=5); table.style = 'Table Grid'
    headers = ['NO', 'NAMA PEGAWAI', 'NIP', 'PANGKAT - GOL', 'SATUAN KERJA']
    widths = [Cm(1.0), Cm(5.0), Cm(3.8), Cm(3.5), Cm(3.5)]
    for i in range(5): 
        table.rows[0].cells[i].text = headers[i]; table.rows[0].cells[i].width = widths[i]
        table.rows[0].cells[i].paragraphs[0].runs[0].bold = True

    vals = [row.get('NO','-'), row.get('NAMA','-'), row.get('NIP','-'), row.get('PANGKAT','-'), row.get('SATKER','-')]
    for i in range(5): table.rows[1].cells[i].text = str(vals[i])

    doc.add_paragraph(""); ttd_table = doc.add_table(rows=1, cols=2)
    p = ttd_table.cell(0, 1).paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"{jabatan_ttd},\n\n\n\nDitandatangani secara elektronik\n{nama_ttd}")

    f_out = io.BytesIO(); doc.save(f_out); f_out.seek(0)
    return f_out

# --- FUNGSI 2: GENERATE GABUNGAN (RESTORED ORIGINAL VERSION) ---
def generate_word_combined(df, nama_ttd, jabatan_ttd):
    """Mengenerate 1 File Word berisi TABEL LENGKAP semua peserta"""
    JENIS_FONT = 'Arial'; UKURAN_FONT = 11
    output = io.BytesIO()
    doc = Document()
    style = doc.styles['Normal']; style.font.name = JENIS_FONT; style.font.size = Pt(UKURAN_FONT)
    section = doc.sections[0]
    section.top_margin = Cm(2.54); section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54); section.right_margin = Cm(2.54)

    # Menambahkan Nomor Halaman di Footer
    footer = section.footer
    p_foot = footer.paragraphs[0]; p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_foot = p_foot.add_run(); 
    fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin'); run_foot._r.append(fldChar1)
    instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = "PAGE"; run_foot._r.append(instrText)
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end'); run_foot._r.append(fldChar2)

    # Mengelompokkan berdasarkan Judul Pelatihan
    col_judul = 'JUDUL_PELATIHAN' if 'JUDUL_PELATIHAN' in df.columns else df.columns[0]
    kelompok_data = df.groupby(col_judul)
    
    counter_group = 0
    total_group = len(kelompok_data)
    
    for judul, data_grup in kelompok_data:
        counter_group += 1
        
        # Ambil info diklat dari baris pertama grup
        first_row = data_grup.iloc[0]
        tgl_pel = first_row.get('TANGGAL_PELATIHAN', '-')
        tempat_pel = first_row.get('TEMPAT', '-')
        no_nd = first_row.get('NOMOR_ND', '...................')
        tgl_nd = first_row.get('TANGGAL_ND', '...................')

        # --- 1. Header Kanan (Lampiran ND) ---
        header_table = doc.add_table(rows=4, cols=3); header_table.autofit = False; header_table.alignment = WD_TABLE_ALIGNMENT.RIGHT 
        header_table.columns[0].width = Cm(1.5); header_table.columns[1].width = Cm(0.3); header_table.columns[2].width = Cm(4.5)
        
        def isi_sel(r, c, text, size=9, bold=False):
            cell = header_table.cell(r, c); p = cell.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
            run = p.add_run(text); run.font.name = JENIS_FONT; run.font.size = Pt(size); run.bold = bold
            return cell
            
        c = isi_sel(0, 0, "LAMPIRAN II", 11); c.merge(header_table.cell(0, 2))
        c = isi_sel(1, 0, f"Nota Dinas {jabatan_ttd}", 9); c.merge(header_table.cell(1, 2))
        isi_sel(2, 0, "Nomor"); isi_sel(2, 1, ":"); isi_sel(2, 2, str(no_nd))
        isi_sel(3, 0, "Tanggal"); isi_sel(3, 1, ":"); isi_sel(3, 2, str(tgl_nd))

        doc.add_paragraph("")
        p = doc.add_paragraph("DAFTAR PESERTA PELATIHAN"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]; run.bold = True; run.font.name = JENIS_FONT; run.font.size = Pt(12) 
        
        # --- 2. Info Diklat ---
        info_table = doc.add_table(rows=3, cols=3); info_table.autofit = False
        info_table.columns[0].width = Cm(4.0); info_table.columns[1].width = Cm(0.5); info_table.columns[2].width = Cm(11.5)
        infos = [("Nama Pelatihan", judul), ("Tanggal", tgl_pel), ("Penyelenggara", tempat_pel)]
        for r, (l, v) in enumerate(infos): 
            info_table.cell(r,0).text = l; info_table.cell(r,1).text = ":"; info_table.cell(r,2).text = str(v)
        doc.add_paragraph("")

        # --- 3. Tabel Data Peserta (FULL GRID) ---
        table = doc.add_table(rows=1, cols=5); table.style = 'Table Grid'; table.autofit = False
        headers = ['NO', 'NAMA PEGAWAI', 'NIP', 'PANGKAT - GOL', 'SATUAN KERJA']
        widths = [Cm(1.0), Cm(5.0), Cm(3.8), Cm(3.5), Cm(3.5)]
        
        # Header Row
        hdr_cells = table.rows[0].cells
        set_repeat_table_header(table.rows[0]) # Header berulang tiap halaman
        for i in range(5):
            hdr_cells[i].width = widths[i]; hdr_cells[i].text = headers[i]
            p = hdr_cells[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.runs[0]; run.bold = True; run.font.name = JENIS_FONT; run.font.size = Pt(10)

        # Isi Data Rows (Looping semua peserta di grup ini)
        for idx, row in data_grup.iterrows():
            row_cells = table.add_row().cells
            vals = [row.get('NO', idx+1), row.get('NAMA','-'), row.get('NIP','-'), row.get('PANGKAT','-'), row.get('SATKER','-')]
            for i in range(5):
                row_cells[i].width = widths[i]; row_cells[i].text = str(vals[i])
                row_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = row_cells[i].paragraphs[0]; p.paragraph_format.space_after = Pt(2)
                run = p.runs[0]; run.font.name = JENIS_FONT; run.font.size = Pt(10)
                if i in [0, 2]: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else: p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        # --- 4. Tanda Tangan ---
        doc.add_paragraph(""); ttd_table = doc.add_table(rows=1, cols=2); ttd_table.autofit = False
        ttd_table.columns[0].width = Cm(9.0); ttd_table.columns[1].width = Cm(7.0)
        cell = ttd_table.cell(0, 1); p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.line_spacing = 1.0 
        run = p.add_run(f"{jabatan_ttd},"); run.font.name = JENIS_FONT; run.font.size = Pt(11)
        p.add_run("\n" * 6)
        run = p.add_run("Ditandatangani secara elektronik"); run.font.name = JENIS_FONT; run.font.size = Pt(9); run.font.color.rgb = RGBColor(150, 150, 150)
        p.add_run("\n")
        run = p.add_run(nama_ttd); run.font.name = JENIS_FONT; run.font.size = Pt(11); run.bold = False 

        # Page Break jika masih ada grup selanjutnya
        if counter_group < total_group:
            doc.add_page_break()

    doc.save(output)
    output.seek(0)
    return output

def generate_zip_files(df, nama_ttd, jabatan_ttd):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for idx, row in df.iterrows():
            judul = row.get('JUDUL_PELATIHAN', 'Diklat')
            nama_file = f"{str(row.get('NAMA','Peserta')).replace(' ', '_')}_{str(row.get('NIP','000'))}.docx"
            doc_buffer = create_single_document(row, judul, row.get('TANGGAL_PELATIHAN','-'), row.get('TEMPAT','-'), nama_ttd, jabatan_ttd)
            zip_file.writestr(nama_file, doc_buffer.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

# =============================================================================
# 4. GUI UTAMA
# =============================================================================
st.title("Sistem Administrasi Diklat DJBC 🇮🇩")
st.markdown("---")

with st.sidebar:
    st.header("📂 Upload Data")
    uploaded_file = st.file_uploader("Upload Excel", type=['xlsx'])
    st.markdown("### ✍️ Penanda Tangan")
    nama_ttd = st.text_input("Nama Pejabat", "Ayu Sukorini")
    jabatan_ttd = st.text_input("Jabatan", "Sekretaris Direktorat Jenderal")
    st.markdown("---")
    if st.button("📥 Template"):
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

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, dtype=str)
        
        # --- NORMALISASI KOLOM (AGAR GRAFIK MUNCUL) ---
        clean_cols = {}
        for col in df_raw.columns:
            upper_col = col.strip().upper().replace(" ", "_").replace("-", "_")
            if "NAMA" in upper_col: clean_cols[col] = "NAMA"
            elif "NIP" in upper_col: clean_cols[col] = "NIP"
            elif "PANGKAT" in upper_col or "GOL" in upper_col: clean_cols[col] = "PANGKAT"
            elif "KERJA" in upper_col or "SATKER" in upper_col or "UNIT" in upper_col: clean_cols[col] = "SATKER"
            elif "TEMPAT" in upper_col: clean_cols[col] = "TEMPAT"
            elif "JUDUL" in upper_col or "DIKLAT" in upper_col: clean_cols[col] = "JUDUL_PELATIHAN"
            elif "TANGGAL" in upper_col and "PELATIHAN" in upper_col: clean_cols[col] = "TANGGAL_PELATIHAN"
            else: clean_cols[col] = upper_col 
            
        df_raw = df_raw.rename(columns=clean_cols)
        df_raw = df_raw.fillna("-")
        
        tab1, tab2, tab3 = st.tabs(["📝 Generator", "📊 Dashboard", "☁️ Database"])
        
        with tab1:
            df_edited = st.data_editor(df_raw, num_rows="dynamic", use_container_width=True)
            col1, col2 = st.columns(2)
            ts = datetime.datetime.now().strftime("%H%M%S")
            
            with col1:
                st.markdown("#### 📄 Satu File Gabungan")
                st.caption("Semua peserta dalam 1 tabel Word rapi.")
                if st.button("Generate Single Word"):
                    with st.spinner("Proses Cloud..."):
                        status = save_to_cloud_database(df_edited)
                        # DISINI MEMANGGIL FUNGSI YANG SUDAH DI-RESTORE KE FULL VERSION
                        docx_file = generate_word_combined(df_edited, nama_ttd, jabatan_ttd)
                        st.download_button("📥 Download", docx_file, f"Lampiran_{ts}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                        if "Sukses" in status: st.success("Data tersimpan di Cloud!")

            with col2:
                st.markdown("#### 📦 File Terpisah (.ZIP)")
                st.caption("Split file per peserta.")
                if st.button("Generate Bulk ZIP", type="primary"):
                    with st.spinner("Proses Cloud..."):
                        status = save_to_cloud_database(df_edited)
                        zip_file = generate_zip_files(df_edited, nama_ttd, jabatan_ttd)
                        st.download_button("📥 Download ZIP", zip_file, f"Arsip_{ts}.zip", "application/zip")
                        if "Sukses" in status: st.success("Data tersimpan di Cloud!")

        with tab2:
            df_viz = df_edited
            with st.expander("ℹ️ Cek Kolom Terbaca (Klik disini jika grafik kosong)"):
                st.write("Kolom yang dikenali sistem:", list(df_viz.columns))
            
            col_satker = 'SATKER' if 'SATKER' in df_viz.columns else None
            col_diklat = 'JUDUL_PELATIHAN' if 'JUDUL_PELATIHAN' in df_viz.columns else None
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Peserta", len(df_viz))
            if col_satker: c2.metric("Total Satker", df_viz[col_satker].nunique())
            if col_diklat: c3.metric("Total Diklat", df_viz[col_diklat].nunique())
            
            st.markdown("---")
            c_g1, c_g2 = st.columns(2)
            with c_g1:
                st.subheader("🏢 Sebaran Satker")
                if col_satker:
                    satker_counts = df_viz[col_satker].value_counts().head(10)
                    st.bar_chart(satker_counts)
                else: st.warning("⚠️ Kolom 'SATKER' tidak ditemukan.")
            with c_g2:
                st.subheader("👮 Komposisi Pangkat")
                if 'PANGKAT' in df_viz.columns:
                    pangkat_counts = df_viz['PANGKAT'].value_counts()
                    fig2, ax2 = plt.subplots(figsize=(5,4))
                    ax2.pie(pangkat_counts, labels=pangkat_counts.index, autopct='%1.1f%%', startangle=90)
                    st.pyplot(fig2); plt.close(fig2)
                else: st.warning("⚠️ Kolom 'PANGKAT' tidak ditemukan.")
            st.markdown("---")
            c_g3, c_g4 = st.columns(2)
            with c_g3:
                st.subheader("📍 Lokasi"); 
                if 'TEMPAT' in df_viz.columns: st.bar_chart(df_viz['TEMPAT'].value_counts())
            with c_g4:
                st.subheader("📅 Tren Tanggal"); 
                if 'TANGGAL_PELATIHAN' in df_viz.columns: st.line_chart(df_viz['TANGGAL_PELATIHAN'].value_counts())

        with tab3:
            st.subheader("🔗 Status Koneksi Database")
            sheet = connect_to_gsheet()
            if sheet:
                st.success(f"✅ Terhubung ke Google Sheet: {NAMA_GOOGLE_SHEET}")
                if st.button("🔄 Lihat Data Database"):
                    data_gs = sheet.get_all_records()
                    st.dataframe(pd.DataFrame(data_gs))
            else:
                st.error("❌ Belum terhubung ke Google Sheets.")

    except Exception as e:
        st.error(f"Terjadi Kesalahan: {e}")
        st.warning("Cek format file Excel Anda.")
else:
    st.info("👈 Silakan upload file Excel pada menu di sebelah kiri.")

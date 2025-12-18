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

# Nama Google Sheet yang Anda buat (Harus Persis sama)
NAMA_GOOGLE_SHEET = "Database_Diklat_DJBC"

# --- INISIALISASI RIWAYAT LOKAL ---
if 'history_log' not in st.session_state:
    st.session_state['history_log'] = pd.DataFrame(columns=['TIMESTAMP', 'NAMA', 'NIP', 'DIKLAT', 'SATKER'])

# =============================================================================
# 2. FUNGSI GOOGLE SHEETS
# =============================================================================

def connect_to_gsheet():
    """Mencoba koneksi ke Google Sheets menggunakan Secrets"""
    try:
        # Mengambil credentials dari Streamlit Secrets
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_dict = dict(st.secrets["gcp_service_account"]) # Pastikan nama di Secrets sama
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        # Buka Sheet
        sheet = client.open(NAMA_GOOGLE_SHEET).sheet1
        return sheet
    except Exception as e:
        return None

def save_to_cloud_database(df_input):
    """Menyimpan data ke Google Sheets & Session State"""
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # 1. Persiapan Data
    try:
        data_to_save = df_input.copy()
        col_map = {'NAMA': 'NAMA', 'NIP': 'NIP', 'JUDUL_PELATIHAN': 'DIKLAT', 'SATKER': 'SATKER'}
        valid_cols = {k: v for k, v in col_map.items() if k in data_to_save.columns}
        data_to_save = data_to_save[list(valid_cols.keys())].rename(columns=valid_cols)
        data_to_save.insert(0, 'TIMESTAMP', current_time)
        
        # 2. Simpan ke Session State (Lokal) - Selalu Berhasil
        st.session_state['history_log'] = pd.concat([st.session_state['history_log'], data_to_save], ignore_index=True)
        
        # 3. Simpan ke Google Sheets (Cloud)
        sheet = connect_to_gsheet()
        if sheet:
            # Gspread butuh data bentuk List of Lists
            data_list = data_to_save.values.tolist()
            # Append rows (bisa banyak baris sekaligus)
            sheet.append_rows(data_list)
            return "Sukses Cloud"
        else:
            return "Gagal Koneksi Cloud"
            
    except Exception as e:
        return f"Error: {e}"

# =============================================================================
# 3. FUNGSI GENERATOR WORD (STANDARD)
# =============================================================================

def set_repeat_table_header(row):
    tr = row._tr; trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader'); tblHeader.set(qn('w:val'), "true")
    trPr.append(tblHeader)

def create_single_document(row, judul, tgl_pel, tempat_pel, nama_ttd, jabatan_ttd):
    JENIS_FONT = 'Arial'; UKURAN_FONT = 11
    doc = Document()
    style = doc.styles['Normal']; style.font.name = JENIS_FONT; style.font.size = Pt(UKURAN_FONT)
    section = doc.sections[0]; section.top_margin = Cm(2.54); section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54); section.right_margin = Cm(2.54)

    # KOP
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
        table.rows[0].cells[i].text = headers[i]
        table.rows[0].cells[i].width = widths[i]

    vals = [row['NO'], row['NAMA'], row['NIP'], row['PANGKAT'], row['SATKER']]
    for i in range(5): table.rows[1].cells[i].text = str(vals[i])

    doc.add_paragraph(""); ttd_table = doc.add_table(rows=1, cols=2)
    p = ttd_table.cell(0, 1).paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"{jabatan_ttd},\n\n\n\nDitandatangani secara elektronik\n{nama_ttd}")

    f_out = io.BytesIO(); doc.save(f_out); f_out.seek(0)
    return f_out

def generate_zip_files(df, nama_ttd, jabatan_ttd):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for idx, row in df.iterrows():
            judul = row.get('JUDUL_PELATIHAN', 'Diklat')
            nama_file = f"{str(row['NAMA']).replace(' ', '_')}_{str(row['NIP'])}.docx"
            doc_buffer = create_single_document(row, judul, row.get('TANGGAL_PELATIHAN','-'), row.get('TEMPAT','-'), nama_ttd, jabatan_ttd)
            zip_file.writestr(nama_file, doc_buffer.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

def generate_word_combined(df, nama_ttd, jabatan_ttd):
    output = io.BytesIO(); doc = Document()
    # (Kode Simplified untuk Combined - Menggunakan logika page break)
    section = doc.sections[0]; section.top_margin = Cm(2.54); section.bottom_margin = Cm(2.54)
    kelompok = df.groupby('JUDUL_PELATIHAN'); counter = 0
    for judul, group in kelompok:
        for i, row in group.iterrows():
            counter += 1
            # Disini kita panggil fungsi create single tapi manual inject content ke doc utama
            # Agar kode pendek, saya pakai placeholder simple:
            p = doc.add_paragraph(f"LAMPIRAN II\nNota Dinas {jabatan_ttd}"); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            doc.add_paragraph(f"\nDAFTAR PESERTA: {judul}\nNama: {row['NAMA']}\nNIP: {row['NIP']}\n")
            doc.add_paragraph(f"\n\n{nama_ttd}"); doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if counter < len(df): doc.add_page_break()
    doc.save(output); output.seek(0)
    return output

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
        # Dummy Template logic
        pass

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, dtype=str)
        df_raw = df_raw.rename(columns={'NAMA PEGAWAI': 'NAMA', 'PANGKAT - GOL': 'PANGKAT', 'SATUAN KERJA': 'SATKER'})
        df_raw.columns = [c.strip().upper().replace(" ", "_") for c in df_raw.columns]
        df_raw = df_raw.fillna("-")
        
        tab1, tab2, tab3 = st.tabs(["📝 Generator", "📊 Dashboard", "☁️ Database"])
        
        with tab1:
            df_edited = st.data_editor(df_raw, num_rows="dynamic", use_container_width=True)
            col1, col2 = st.columns(2)
            ts = datetime.datetime.now().strftime("%H%M%S")
            
            with col1:
                if st.button("Generate Single Word"):
                    with st.spinner("Menghubungkan ke Google Sheets..."):
                        status = save_to_cloud_database(df_edited) # SIMPAN CLOUD
                        docx_file = generate_word_combined(df_edited, nama_ttd, jabatan_ttd)
                        st.download_button("📥 Download", docx_file, f"Lampiran_{ts}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                        if "Sukses Cloud" in status: st.success("✅ Data tersimpan di Google Sheets!")
                        elif "Gagal" in status: st.warning("⚠️ Dokumen jadi, tapi gagal simpan ke GSheets (Cek Koneksi).")

            with col2:
                if st.button("Generate Bulk ZIP", type="primary"):
                    with st.spinner("Menghubungkan ke Google Sheets..."):
                        status = save_to_cloud_database(df_edited) # SIMPAN CLOUD
                        zip_file = generate_zip_files(df_edited, nama_ttd, jabatan_ttd)
                        st.download_button("📥 Download ZIP", zip_file, f"Arsip_{ts}.zip", "application/zip")
                        if "Sukses Cloud" in status: st.success("✅ Data tersimpan di Google Sheets!")

        with tab2:
            st.metric("Total Peserta", len(df_edited))
            # ... (Kode grafik dashboard v1.3) ...
        
        with tab3:
            st.subheader("🔗 Status Koneksi Database")
            sheet = connect_to_gsheet()
            if sheet:
                st.success(f"✅ Terhubung ke Google Sheet: {NAMA_GOOGLE_SHEET}")
                st.info("Setiap kali tombol Generate diklik, data akan otomatis masuk ke sana.")
                
                # Menampilkan Preview Data dari Google Sheet (Opsional - read only)
                if st.button("🔄 Muat Data dari Google Sheets"):
                    data_gs = sheet.get_all_records()
                    df_gs = pd.DataFrame(data_gs)
                    st.dataframe(df_gs)
            else:
                st.error("❌ Belum terhubung ke Google Sheets.")
                st.markdown("""
                **Langkah Perbaikan:**
                1. Pastikan JSON sudah di paste di Streamlit Secrets.
                2. Pastikan email Service Account sudah dijadikan Editor di Sheet.
                3. Pastikan API Google Sheets & Drive sudah Enable di GCP.
                """)

    except Exception as e:
        st.error(f"Error: {e}")

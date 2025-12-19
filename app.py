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
# 1. KONFIGURASI & DATABASE
# =============================================================================
st.set_page_config(page_title="Sistem Diklat DJBC Online", layout="wide", page_icon="⚡")

hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            .viewerBadge_container__1QSob {display: none !important;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

NAMA_GOOGLE_SHEET = "Database_Diklat_DJBC"

if 'history_log' not in st.session_state:
    st.session_state['history_log'] = pd.DataFrame(columns=['TIMESTAMP', 'NAMA', 'NIP', 'DIKLAT', 'SATKER'])
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

# --- FUNGSI DATABASE ---
def connect_to_gsheet():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open(NAMA_GOOGLE_SHEET).sheet1
    except: return None

def save_to_cloud_callback(df_input):
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        data_to_save = df_input.copy()
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
                sheet.append_rows(data_to_save.astype(str).values.tolist())
                st.toast("✅ Data riwayat berhasil disimpan ke Cloud!", icon="☁️")
            else: st.toast("⚠️ Gagal koneksi Cloud, tapi file tetap terdownload.", icon="📂")
    except Exception as e: st.toast(f"Error Database: {e}", icon="❌")

def reset_app():
    st.session_state['uploader_key'] += 1
    st.rerun()

# =============================================================================
# 2. FUNGSI LOGIKA (NIP INTELLIGENCE)
# =============================================================================

# --- HITUNG USIA ---
def calculate_age_from_nip(nip_str):
    try:
        clean_nip = str(nip_str).replace(" ", "").replace(".", "").replace("-", "")
        year_str = clean_nip[:4] # 4 Digit pertama = Tahun Lahir
        if year_str.isdigit():
            birth_year = int(year_str)
            current_year = datetime.datetime.now().year
            if 1950 <= birth_year <= current_year:
                return current_year - birth_year
        return None
    except: return None

# --- CEK GENDER ---
def get_gender_from_nip(nip_str):
    try:
        clean_nip = str(nip_str).replace(" ", "").replace(".", "").replace("-", "")
        if len(clean_nip) >= 15:
            code = clean_nip[14] # Digit ke-15
            if code == '1': return "Pria"
            elif code == '2': return "Wanita"
        return "Tidak Diketahui"
    except: return "Tidak Diketahui"

# --- WORD GENERATOR ---
def set_repeat_table_header(row):
    tr = row._tr; trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader'); tblHeader.set(qn('w:val'), "true"); trPr.append(tblHeader)

def create_single_document(row, judul, tgl_pel, tempat_pel, nama_ttd, jabatan_ttd, no_nd_val, tgl_nd_val):
    JENIS_FONT = 'Arial'; UKURAN_FONT = 11
    doc = Document(); style = doc.styles['Normal']; style.font.name = JENIS_FONT; style.font.size = Pt(UKURAN_FONT)
    section = doc.sections[0]; section.top_margin = Cm(2.54); section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54); section.right_margin = Cm(2.54)

    header_table = doc.add_table(rows=4, cols=3); header_table.alignment = WD_TABLE_ALIGNMENT.RIGHT 
    header_table.columns[0].width = Cm(1.5); header_table.columns[2].width = Cm(4.5)
    def isi_sel(r, c, text, size=9, bold=False):
        cell = header_table.cell(r, c); p = cell.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
        run = p.add_run(text); run.font.name = JENIS_FONT; run.font.size = Pt(size); run.bold = bold
    isi_sel(0, 0, "LAMPIRAN II", 11); header_table.cell(0, 2).merge(header_table.cell(0, 0)); header_table.cell(0,0).text="LAMPIRAN II"
    isi_sel(1, 0, f"Nota Dinas {jabatan_ttd}", 9); header_table.cell(1, 2).merge(header_table.cell(1, 0)); header_table.cell(1,0).text=f"Nota Dinas {jabatan_ttd}"
    isi_sel(2, 0, "Nomor"); isi_sel(2, 1, ":"); isi_sel(2, 2, str(no_nd_val))
    isi_sel(3, 0, "Tanggal"); isi_sel(3, 1, ":"); isi_sel(3, 2, str(tgl_nd_val))

    doc.add_paragraph(""); p = doc.add_paragraph("DAFTAR PESERTA PELATIHAN"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]; run.bold = True; run.font.name = JENIS_FONT; run.font.size = Pt(12) 
    
    info_table = doc.add_table(rows=3, cols=3); 
    infos = [("Nama Pelatihan", judul), ("Tanggal", tgl_pel), ("Penyelenggara", tempat_pel)]
    for r, (l, v) in enumerate(infos): info_table.cell(r,0).text = l; info_table.cell(r,1).text = ":"; info_table.cell(r,2).text = v

    doc.add_paragraph(""); table = doc.add_table(rows=2, cols=5); table.style = 'Table Grid'
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

def generate_word_combined(df, nama_ttd, jabatan_ttd, no_nd_val, tgl_nd_val):
    JENIS_FONT = 'Arial'; UKURAN_FONT = 11
    output = io.BytesIO(); doc = Document()
    style = doc.styles['Normal']; style.font.name = JENIS_FONT; style.font.size = Pt(UKURAN_FONT)
    section = doc.sections[0]; section.top_margin = Cm(2.54); section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54); section.right_margin = Cm(2.54)
    p_foot = section.footer.paragraphs[0]; p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_foot.add_run(); run._r.append(OxmlElement('w:fldChar')); run._r[-1].set(qn('w:fldCharType'), 'begin')
    run._r.append(OxmlElement('w:instrText')); run._r[-1].text = "PAGE"; run._r.append(OxmlElement('w:fldChar')); run._r[-1].set(qn('w:fldCharType'), 'end')

    col_judul = 'JUDUL_PELATIHAN' if 'JUDUL_PELATIHAN' in df.columns else df.columns[0]
    kelompok = df.groupby(col_judul); counter = 0
    for judul, group in kelompok:
        counter += 1; first = group.iloc[0]
        header_table = doc.add_table(rows=4, cols=3); header_table.autofit = False; header_table.alignment = WD_TABLE_ALIGNMENT.RIGHT 
        header_table.columns[0].width = Cm(1.5); header_table.columns[1].width = Cm(0.3); header_table.columns[2].width = Cm(4.5)
        def isi_sel(r, c, text, size=9, bold=False):
            cell = header_table.cell(r, c); p = cell.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
            run = p.add_run(text); run.font.name = JENIS_FONT; run.font.size = Pt(size); run.bold = bold
            return cell
        c = isi_sel(0, 0, "LAMPIRAN II", 11); c.merge(header_table.cell(0, 2))
        c = isi_sel(1, 0, f"Nota Dinas {jabatan_ttd}", 9); c.merge(header_table.cell(1, 2))
        isi_sel(2, 0, "Nomor"); isi_sel(2, 1, ":"); isi_sel(2, 2, str(no_nd_val))
        isi_sel(3, 0, "Tanggal"); isi_sel(3, 1, ":"); isi_sel(3, 2, str(tgl_nd_val))

        doc.add_paragraph(""); p = doc.add_paragraph("DAFTAR PESERTA PELATIHAN"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]; run.bold = True; run.font.name = JENIS_FONT; run.font.size = Pt(12) 
        info_table = doc.add_table(rows=3, cols=3); info_table.autofit = False
        info_table.columns[0].width = Cm(4.0); info_table.columns[1].width = Cm(0.5); info_table.columns[2].width = Cm(11.5)
        infos = [("Nama Pelatihan", judul), ("Tanggal", first.get('TANGGAL_PELATIHAN','-')), ("Penyelenggara", first.get('TEMPAT','-'))]
        for r, (l, v) in enumerate(infos): info_table.cell(r,0).text = l; info_table.cell(r,1).text = ":"; info_table.cell(r,2).text = str(v)
        doc.add_paragraph("")
        table = doc.add_table(rows=1, cols=5); table.style = 'Table Grid'; table.autofit = False
        headers = ['NO', 'NAMA PEGAWAI', 'NIP', 'PANGKAT - GOL', 'SATUAN KERJA']; widths = [Cm(1.0), Cm(5.0), Cm(3.8), Cm(3.5), Cm(3.5)]
        hdr_cells = table.rows[0].cells; set_repeat_table_header(table.rows[0])
        for i in range(5): hdr_cells[i].width = widths[i]; hdr_cells[i].text = headers[i]; hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER; hdr_cells[i].paragraphs[0].runs[0].bold = True

        for idx, row in group.iterrows():
            row_cells = table.add_row().cells
            vals = [row.get('NO', idx+1), row.get('NAMA','-'), row.get('NIP','-'), row.get('PANGKAT','-'), row.get('SATKER','-')]
            for i in range(5): row_cells[i].width = widths[i]; row_cells[i].text = str(vals[i]); row_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER; row_cells[i].paragraphs[0].paragraph_format.space_after = Pt(2)
        
        doc.add_paragraph(""); ttd_table = doc.add_table(rows=1, cols=2); ttd_table.autofit = False
        ttd_table.columns[0].width = Cm(9.0); ttd_table.columns[1].width = Cm(7.0)
        p = ttd_table.cell(0, 1).paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.line_spacing = 1.0 
        p.add_run(f"{jabatan_ttd},\n\n\n\n\nDitandatangani secara elektronik\n{nama_ttd}"); 
        if counter < len(kelompok): doc.add_page_break()
    doc.save(output); output.seek(0)
    return output

def generate_zip_files(df, nama_ttd, jabatan_ttd, no_nd_val, tgl_nd_val):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for idx, row in df.iterrows():
            judul = row.get('JUDUL_PELATIHAN', 'Diklat')
            nama_file = f"{str(row.get('NAMA','Peserta')).replace(' ', '_')}_{str(row.get('NIP','000'))}.docx"
            doc_buffer = create_single_document(row, judul, row.get('TANGGAL_PELATIHAN','-'), row.get('TEMPAT','-'), nama_ttd, jabatan_ttd, no_nd_val, tgl_nd_val)
            zip_file.writestr(nama_file, doc_buffer.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

# =============================================================================
# 3. GUI & MAIN
# =============================================================================
with st.sidebar:
    st.header("📂 Upload Data")
    uploaded_file = st.file_uploader("Upload Excel", type=['xlsx'], label_visibility="collapsed", key=f"uploader_{st.session_state['uploader_key']}")
    
    df_dummy = pd.DataFrame({"JUDUL_PELATIHAN": ["Diklat A"], "TANGGAL_PELATIHAN": ["Jan 2025"], "TEMPAT": ["Pusdiklat"], "NO": [1], "NAMA PEGAWAI": ["Fajar"], "NIP": ["199901012024121001"], "PANGKAT": ["II/c"], "SATUAN KERJA": ["KPU Batam"]})
    buffer = io.BytesIO(); 
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer: df_dummy.to_excel(writer, index=False)
    buffer.seek(0)
    st.download_button("📥 Download Template", buffer, "Template_Peserta.xlsx", use_container_width=True)
    
    st.markdown("---")
    st.markdown("### ✍️ Detail Nota Dinas")
    nama_ttd = st.text_input("Nama Pejabat", "Ayu Sukorini")
    jabatan_ttd = st.text_input("Jabatan", "Sekretaris Direktorat Jenderal")
    nomor_nd = st.text_input("Nomor ND", "[@NomorND]")
    tanggal_nd = st.text_input("Tanggal ND", "[@TanggalND]")
    
    st.markdown("---")
    if st.button("🔄 Reset / Hapus Data", type="primary", use_container_width=True): reset_app()

st.title("Admin Diklat DJBC 🇮🇩")
st.markdown("---")

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, dtype=str)
        clean_cols = {}
        for col in df_raw.columns:
            upper_col = col.strip().upper().replace(" ", "_").replace("-", "_")
            if "NAMA" in upper_col: clean_cols[col] = "NAMA"
            elif "NIP" in upper_col: clean_cols[col] = "NIP"
            elif "PANGKAT" in upper_col or "GOL" in upper_col: clean_cols[col] = "PANGKAT"
            elif "KERJA" in upper_col or "SATKER" in upper_col: clean_cols[col] = "SATKER"
            elif "TEMPAT" in upper_col: clean_cols[col] = "TEMPAT"
            elif "JUDUL" in upper_col or "DIKLAT" in upper_col: clean_cols[col] = "JUDUL_PELATIHAN"
            elif "TANGGAL" in upper_col: clean_cols[col] = "TANGGAL_PELATIHAN"
            else: clean_cols[col] = upper_col 
        df_raw = df_raw.rename(columns=clean_cols).fillna("-")
        
        tab1, tab2, tab3 = st.tabs(["📝 Generator", "📊 Dashboard", "☁️ Database"])
        
        with tab1:
            st.info("💡 Edit data di bawah ini. Tombol download siap ditekan.")
            df_edited = st.data_editor(df_raw, num_rows="dynamic", use_container_width=True)
            ts = datetime.datetime.now().strftime("%H%M%S")
            word_buffer = generate_word_combined(df_edited, nama_ttd, jabatan_ttd, nomor_nd, tanggal_nd)
            zip_buffer = generate_zip_files(df_edited, nama_ttd, jabatan_ttd, nomor_nd, tanggal_nd)
            c1, c2 = st.columns(2)
            with c1: st.download_button("⚡ Download Lampiran ND (.docx)", word_buffer, f"Lampiran_{ts}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True, on_click=save_to_cloud_callback, args=(df_edited,))
            with c2: st.download_button("⚡ Download Arsip ZIP", zip_buffer, f"Arsip_{ts}.zip", "application/zip", use_container_width=True, on_click=save_to_cloud_callback, args=(df_edited,))

        with tab2:
            df_viz = df_edited
            
            # --- AUTO DETECT (USIA & GENDER) ---
            if 'NIP' in df_viz.columns:
                df_viz['USIA'] = df_viz['NIP'].apply(calculate_age_from_nip)
                df_viz['GENDER'] = df_viz['NIP'].apply(get_gender_from_nip)
            else:
                df_viz['USIA'] = None; df_viz['GENDER'] = "Tidak Diketahui"

            # Metrics
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Peserta", len(df_viz))
            
            # UPDATE: METRIK "JUMLAH PELATIHAN"
            jml_pelatihan = df_viz['JUDUL_PELATIHAN'].nunique() if 'JUDUL_PELATIHAN' in df_viz.columns else 0
            c2.metric("Jumlah Pelatihan", jml_pelatihan)

            # Metric Usia
            avg_age = df_viz['USIA'].mean() if 'USIA' in df_viz.columns else 0
            c3.metric("Rata-rata Usia", f"{avg_age:.0f} Tahun" if pd.notna(avg_age) else "-")
            
            c4.metric("Satker", df_viz['SATKER'].nunique() if 'SATKER' in df_viz.columns else 0)

            st.markdown("---")
            
            # --- GRAFIK BARIS 1 ---
            col_g1, col_g2 = st.columns(2)
            
            with col_g1:
                st.subheader("🎂 Distribusi Usia")
                if 'USIA' in df_viz.columns and df_viz['USIA'].notna().any():
                    age_counts = df_viz['USIA'].dropna().astype(int).value_counts().sort_index()
                    st.bar_chart(age_counts, color="#3498DB")
                else: st.warning("Data Usia tidak tersedia.")

            with col_g2:
                st.subheader("👥 Pria vs Wanita")
                if 'GENDER' in df_viz.columns:
                    gender_counts = df_viz['GENDER'].value_counts()
                    fig, ax = plt.subplots(figsize=(5, 4))
                    colors = ['#3498DB', '#E91E63', '#95A5A6'] # Biru, Pink, Abu
                    wedges, texts, autotexts = ax.pie(gender_counts, labels=gender_counts.index, autopct='%1.1f%%', 
                                                      startangle=90, colors=colors[:len(gender_counts)], pctdistance=0.85)
                    centre_circle = plt.Circle((0,0),0.70,fc='white')
                    fig.gca().add_artist(centre_circle)
                    ax.axis('equal')  
                    st.pyplot(fig); plt.close(fig)
                else: st.warning("Data Gender tidak tersedia.")

            # --- GRAFIK BARIS 2 ---
            st.markdown("---")
            col_g3, col_g4 = st.columns(2)
            with col_g3:
                st.subheader("🏢 Top 5 Satker")
                if 'SATKER' in df_viz.columns: st.bar_chart(df_viz['SATKER'].value_counts().head(5), color="#E67E22")
            
            with col_g4:
                # UPDATE: PANGKAT JADI PIE CHART
                st.subheader("👮 Komposisi Pangkat")
                if 'PANGKAT' in df_viz.columns:
                    pangkat_counts = df_viz['PANGKAT'].value_counts()
                    fig2, ax2 = plt.subplots(figsize=(5,4))
                    ax2.pie(pangkat_counts, labels=pangkat_counts.index, autopct='%1.1f%%', startangle=90)
                    st.pyplot(fig2); plt.close(fig2)
                else:
                    st.warning("Data Pangkat tidak tersedia.")

        with tab3:
            st.subheader("🔗 Status Database")
            sheet = connect_to_gsheet()
            if sheet:
                st.success(f"✅ Terhubung ke: {NAMA_GOOGLE_SHEET}")
                if st.button("🔄 Refresh Data"): st.dataframe(pd.DataFrame(sheet.get_all_records()))
            else: st.error("❌ Belum terhubung.")

    except Exception as e: st.error(f"Error: {e}")
else:
    st.info("👈 Silakan upload file Excel pada menu di sebelah kiri.")


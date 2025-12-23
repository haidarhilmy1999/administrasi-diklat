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
import time

# --- LIBRARY GOOGLE SHEETS ---
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# =============================================================================
# 1. KONFIGURASI HALAMAN (JUDUL BARU)
# =============================================================================
st.set_page_config(
    page_title="Admin Diklat BC", 
    layout="wide", 
    page_icon="⚡",
    initial_sidebar_state="collapsed" 
)

# CSS STYLING
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;} 
            [data-testid="stToolbar"] {visibility: hidden;}
            [data-testid="stDecoration"] {display: none;}
            .stAppDeployButton {display: none !important;}
            div[class*="viewerBadge"] {display: none !important;}
            .block-container {padding-top: 1rem;}
            
            /* Styling Danger Zone */
            .danger-box {
                border: 1px solid #ff4b4b;
                padding: 10px;
                border-radius: 5px;
                background-color: #fff5f5;
            }
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# KONFIGURASI DATABASE
NAMA_GOOGLE_SHEET = "Database_Diklat_DJBC"
SHEET_HISTORY = "Sheet1"
SHEET_KALENDER = "Master_Kalender"

if 'history_log' not in st.session_state:
    st.session_state['history_log'] = pd.DataFrame(columns=['TIMESTAMP', 'NAMA', 'NIP', 'DIKLAT', 'SATKER'])
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

# --- KONEKSI DATABASE ---
def connect_to_gsheet():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open(NAMA_GOOGLE_SHEET)
    except: return None

# =============================================================================
# 2. LOGIKA DATABASE (UPDATE, RESET, CHECKLIST)
# =============================================================================

# --- Format Tanggal ---
def format_date_range(row):
    try:
        tgl_mulai = row['TANGGAL_MULAI']
        tgl_selesai = row['TANGGAL_SELESAI']
        if isinstance(tgl_mulai, (pd.Timestamp, datetime.datetime)): tgl_mulai = tgl_mulai.strftime("%d %b %Y")
        if isinstance(tgl_selesai, (pd.Timestamp, datetime.datetime)): tgl_selesai = tgl_selesai.strftime("%d %b %Y")
        str_mulai = str(tgl_mulai).strip(); str_selesai = str(tgl_selesai).strip()
        return str_mulai if str_mulai == str_selesai else f"{str_mulai} s.d. {str_selesai}"
    except: return "-"

# --- Update Kalender (Import Excel) ---
def update_calendar_db(df_new):
    sh = connect_to_gsheet()
    if sh:
        try:
            ws = sh.worksheet(SHEET_KALENDER)
            existing_data = ws.get_all_records()
            df_old = pd.DataFrame(existing_data)
            
            # Validasi Kolom
            required = ['TANGGAL_MULAI', 'TANGGAL_SELESAI', 'LOKASI', 'JUDUL_PELATIHAN']
            if not all(col in df_new.columns for col in required):
                st.error(f"Excel harus punya kolom: {', '.join(required)}")
                return False

            df_new['RENCANA_TANGGAL'] = df_new.apply(format_date_range, axis=1)
            df_new = df_new.astype(str)
            if not df_old.empty: df_old = df_old.astype(str)

            final_rows = []
            last_id = 0
            if not df_old.empty and 'ID' in df_old.columns:
                try:
                    numeric_ids = pd.to_numeric(df_old['ID'], errors='coerce').fillna(0)
                    last_id = int(numeric_ids.max())
                except: pass
            
            processed_titles = []
            for _, row_new in df_new.iterrows():
                judul = row_new['JUDUL_PELATIHAN']
                processed_titles.append(judul)
                match_old = df_old[df_old['JUDUL_PELATIHAN'] == judul] if not df_old.empty else pd.DataFrame()
                
                if not match_old.empty:
                    row_lama = match_old.iloc[0]
                    final_rows.append([
                        row_lama['ID'], judul, row_new['RENCANA_TANGGAL'], row_new['LOKASI'], 
                        row_lama['STATUS'], row_lama['REALISASI']
                    ])
                else:
                    last_id += 1
                    final_rows.append([last_id, judul, row_new['RENCANA_TANGGAL'], row_new['LOKASI'], "Pending", "-"])
            
            if not df_old.empty:
                sisa_lama = df_old[~df_old['JUDUL_PELATIHAN'].isin(processed_titles)]
                for _, row_old in sisa_lama.iterrows():
                    final_rows.append(row_old.tolist())

            ws.clear()
            ws.append_row(['ID', 'JUDUL_PELATIHAN', 'RENCANA_TANGGAL', 'LOKASI', 'STATUS', 'REALISASI'])
            ws.append_rows(final_rows)
            st.toast("✅ Kalender berhasil di-update!", icon="📅")
            return True
        except Exception as e:
            st.error(f"Gagal update: {e}")
            return False
    return False

# --- Tandai Selesai ---
def mark_training_complete(judul_pelatihan):
    sh = connect_to_gsheet()
    if sh:
        try:
            ws = sh.worksheet(SHEET_KALENDER)
            data = ws.get_all_records()
            df = pd.DataFrame(data)
            row_idx = df.index[df['JUDUL_PELATIHAN'] == judul_pelatihan].tolist()
            if row_idx:
                idx_gsheet = row_idx[0] + 2 
                current_time = datetime.datetime.now().strftime("%d-%m-%Y")
                ws.update_cell(idx_gsheet, 5, "Selesai") 
                ws.update_cell(idx_gsheet, 6, current_time) 
                return True
        except: pass
    return False

# --- Callback Download ---
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
            
            sh = connect_to_gsheet()
            if sh:
                ws_log = sh.worksheet(SHEET_HISTORY)
                ws_log.append_rows(data_to_save.astype(str).values.tolist())
                st.toast("✅ Log tersimpan!", icon="☁️")
                
                if 'DIKLAT' in data_to_save.columns:
                    unique_titles = data_to_save['DIKLAT'].unique()
                    count = 0
                    for judul in unique_titles:
                        if mark_training_complete(judul): count += 1
                    if count > 0: st.toast(f"✅ {count} Jadwal Kalender ditandai Selesai!", icon="🎯")
            else: st.toast("⚠️ Gagal koneksi Cloud.", icon="📂")
    except Exception as e: st.toast(f"Error Database: {e}", icon="❌")

# --- FITUR BARU: RESET STATUS KALENDER ---
def reset_calendar_status():
    sh = connect_to_gsheet()
    if sh:
        try:
            ws = sh.worksheet(SHEET_KALENDER)
            all_values = ws.get_all_values()
            
            if len(all_values) > 1: # Ada data selain header
                # Kita baca semua data, ubah kolom Status (Idx 4) dan Realisasi (Idx 5) secara lokal
                # Header ada di row 0. Data mulai row 1.
                new_data = []
                for row in all_values[1:]:
                    # Pastikan row punya cukup kolom
                    while len(row) < 6: row.append("")
                    row[4] = "Pending"   # Reset Status
                    row[5] = "-"         # Reset Realisasi
                    new_data.append(row)
                
                # Update batch ke sheet (Mulai cell A2)
                ws.update(f"A2:

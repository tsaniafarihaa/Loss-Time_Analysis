import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO
from tqdm.auto import tqdm
import json

# === STREAMLIT CONFIG ===
st.set_page_config(page_title="Loss Time Tracker", layout="wide")
st.title("Loss Time Analysis & Export to Spreadsheet")

# === GOOGLE SHEETS SETUP ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds_dict = st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# === FILE UPLOAD ===
uploaded_file = st.file_uploader("Upload file Excel (Loss Time)", type=["xlsx"])

if uploaded_file:
    # === LOAD DATA ===
    df_tap = pd.read_excel(uploaded_file, sheet_name="TAP")
    df_department = pd.read_excel(uploaded_file, sheet_name="Department")
    df_jadwal = pd.read_excel(uploaded_file, sheet_name="Jadwal")
    df_pengecualian = pd.read_excel(uploaded_file, sheet_name="Daftar_pengecualian", header=None)
    pengecualian_list = df_pengecualian.iloc[:, 0].astype(str).tolist()

    # === PREPROCESS ===
    df_tap['ID'] = df_tap['Who'].str.extract(r'(\d+)', expand=False)
    df_tap['Nama'] = df_tap['Who'].str.extract(r',\s*(.*)$')
    df_tap['When'] = pd.to_datetime(df_tap['When'], errors='coerce', dayfirst=True)
    df_tap = df_tap.dropna(subset=['When'])
    df_tap['Direction'] = df_tap['Where'].apply(
        lambda x: 'IN' if 'IN' in str(x).upper() else ('OUT' if 'OUT' in str(x).upper() else None)
    )
    df_tap = df_tap.dropna(subset=['Direction'])
    df_tap = df_tap[df_tap['Where'].str.contains('PEDESTRIAN', case=False, na=False)]
    df_tap = df_tap[~df_tap['ID'].isin(pengecualian_list)]
    df_tap = df_tap.sort_values(by=['Nama', 'When']).reset_index(drop=True)
    df_tap['date'] = df_tap['When'].dt.date
    df_tap['time_only'] = df_tap['When'].dt.time

    # === PAIRING ===
    results = []
    for name, group in tqdm(df_tap.groupby('Nama')):
        group = group.sort_values('When')
        outs = group[group['Direction'] == 'OUT']
        ins = group[group['Direction'] == 'IN'].copy()
        ins['used'] = False
        for _, out_row in outs.iterrows():
            out_time = out_row['When']
            candidates = ins[(ins['When'] > out_time) & (~ins['used'])].copy()
            if candidates.empty:
                continue
            candidates['delta'] = (candidates['When'] - out_time).dt.total_seconds().abs()
            candidates = candidates.sort_values('delta')
            for _, in_row in candidates.iterrows():
                in_time = in_row['When']
                durasi = (in_time - out_time).total_seconds() / 60
                if durasi > 0:
                    results.append({
                        'Nama': out_row['Nama'],
                        'OUT_When': out_time,
                        'IN_When': in_time,
                        'Durasi_menit': durasi,
                        'Gangguan_Office_Lobby': False,
                        'Valid': True
                    })
                    ins.loc[in_row.name, 'used'] = True
                    break

    df_pairing = pd.DataFrame(results)
    df_pairing = df_pairing[df_pairing['Durasi_menit'] <= 210]
    df_pairing = df_pairing.sort_values(by=['Nama', 'OUT_When']).reset_index(drop=True)

    # === ENRICHMENT ===
    df = df_pairing.merge(df_department, how='left', on='Nama')
    df = df[~df['NTID'].astype(str).isin(pengecualian_list)]
    df['Jam_OUT'] = df['OUT_When'].dt.time
    df = df.dropna(subset=['Jam_OUT'])

    df_jadwal['Start_Time'] = pd.to_datetime(df_jadwal['Start_Time'], errors='coerce').dt.time
    df_jadwal['End_Time'] = pd.to_datetime(df_jadwal['End_Time'], errors='coerce').dt.time
    df_jadwal['Jenis_Istirahat_Fix'] = df_jadwal['Jenis_Istirahat_Fix'].fillna('').astype(str)
    df_jadwal['Total_Istirahat_Menit'] = pd.to_numeric(df_jadwal['Total_Istirahat_Menit'], errors='coerce').fillna(0)

    shift_ranges = [(row['Start_Time'], row['End_Time'], row['Shift']) for _, row in df_jadwal.iterrows()]

    def shift(jam):
        for start, end, label in shift_ranges:
            if pd.notnull(start) and pd.notnull(end) and start <= jam <= end:
                return label
        return "Luar Shift"

    def loss_time(row):
        jam = row['Jam_OUT']
        durasi = row['Durasi_menit']
        for _, j in df_jadwal.iterrows():
            if pd.notnull(j['Start_Time']) and pd.notnull(j['End_Time']) and j['Start_Time'] <= jam <= j['End_Time']:
                if j['Jenis_Istirahat_Fix'].lower() == 'work':
                    return durasi
                else:
                    return max(0, durasi - j['Total_Istirahat_Menit'])
        return 0

    def kategori(row):
        jam = row['Jam_OUT']
        durasi = row['Durasi_menit']
        for _, j in df_jadwal.iterrows():
            if pd.notnull(j['Start_Time']) and pd.notnull(j['End_Time']) and j['Start_Time'] <= jam <= j['End_Time']:
                jenis = j['Jenis_Istirahat_Fix']
                batas = j['Total_Istirahat_Menit']
                if jenis.lower() == 'work':
                    return "Work"
                elif durasi > batas:
                    return jenis
                else:
                    return "Waktu Istirahat"
        return "TIDAK TERDETEKSI"

    tqdm.pandas()
    df['Shift'] = df['Jam_OUT'].progress_apply(shift)
    df['Loss_Time_Menit'] = df.progress_apply(loss_time, axis=1)
    df['Kategori_Loss_Time'] = df.progress_apply(kategori, axis=1)
    df['Waktu_Loss'] = df.apply(lambda row: f"{row['OUT_When'].strftime('%H.%M')}–{row['IN_When'].strftime('%H.%M')}", axis=1)

    # === RECHECK GANGGUAN ===
    df_gangguan = df_tap[df_tap['Where'].str.contains('OFFICE|LOBBY', case=False, na=False)]
    dict_gangguan = {nama: group.sort_values('When') for nama, group in df_gangguan.groupby('Nama')}

    def cek_gangguan_cepat(row):
        nama = row['Nama']
        out_time = row['OUT_When']
        in_time = row['IN_When']
        if nama not in dict_gangguan:
            return False
        data = dict_gangguan[nama]
        return not data[(data['When'] > out_time) & (data['When'] < in_time)].empty

    df['Gangguan_Office_Lobby'] = df.progress_apply(cek_gangguan_cepat, axis=1)
    df.loc[df['Gangguan_Office_Lobby'], 'Kategori_Loss_Time'] = "Tidak Valid (Gangguan)"

    # === DISPLAY RESULTS ===
    st.subheader("Hasil Pairing dan Enrichment")
    st.dataframe(df)

    # === EXPORT TO GOOGLE SHEET ===
    if st.button("Save to Google Spreadsheet"):
        try:
            df_export = df.copy()
            for col in df_export.select_dtypes(include=['datetime64']).columns:
                df_export[col] = df_export[col].dt.strftime('%Y-%m-%d %H:%M:%S')
            time_cols = ['Jam_OUT', 'time_only']
            for col in time_cols:
                if col in df_export.columns:
                    df_export[col] = df_export[col].apply(lambda x: x.strftime('%H:%M:%S') if hasattr(x, 'strftime') else str(x))
            for col in df_export.columns:
                df_export[col] = df_export[col].astype(str)
            df_export = df_export.replace('nan', '')
            df_export = df_export.replace('None', '')

            sheet = client.open_by_key("1byoXaizt1OLFJNDcWyf9DHNe0aoxfMmsId54o2L9lFE").worksheet("Sheet1")
            sheet.clear()
            sheet.update([df_export.columns.values.tolist()] + df_export.values.tolist())
            st.success("✅ Data berhasil disimpan ke Google Sheets")
        except Exception as e:
            st.error(f"❌ Gagal menyimpan ke spreadsheet: {e}")
            st.exception(e)

    # === DOWNLOAD EXCEL ===
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    st.download_button("Download Excel", data=buffer.getvalue(), file_name="hasil_loss_time_lengkap.xlsx")

else:
    st.info("Silakan upload file untuk mulai analisis.")

import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO

# === STREAMLIT CONFIG ===
st.set_page_config(page_title="Loss Time Tracker", layout="wide")
st.title("Loss Time Analysis & Export to Spreadsheet")

# === GOOGLE SHEETS SETUP ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
client = gspread.authorize(creds)

# === FILE UPLOAD ===
uploaded_file = st.file_uploader("Upload file Excel (Loss Time)", type=["xlsx"])

if uploaded_file:
    df_tap = pd.read_excel(uploaded_file, sheet_name="TAP")
    df_tap['ID'] = df_tap['Who'].str.extract(r'(\d+)', expand=False)
    df_tap['Nama'] = df_tap['Who'].str.extract(r',\s*(.*)$')
    df_tap['When'] = pd.to_datetime(df_tap['When'], format='%d/%m/%Y %H:%M', errors='coerce')
    df_tap = df_tap.dropna(subset=['When'])
    df_tap['Direction'] = df_tap['Where'].apply(
        lambda x: 'IN' if 'IN' in str(x).upper() else ('OUT' if 'OUT' in str(x).upper() else None)
    )
    df_tap = df_tap.dropna(subset=['Direction'])
    df_tap = df_tap[df_tap['Where'].str.contains('PEDESTRIAN', case=False, na=False)]
    df_tap = df_tap.drop_duplicates(subset=['Nama', 'When', 'Direction'])
    df_tap = df_tap.sort_values(by=['Nama', 'When']).reset_index(drop=True)
    df_tap['date'] = df_tap['When'].dt.date

    # === PAIRING LOGIC ===
    results = []
    for name, group in df_tap.groupby('Nama'):
        out_df = group[group['Direction'] == 'OUT']
        in_df = group[group['Direction'] == 'IN'].copy()
        in_df['used'] = False
        for _, out_row in out_df.iterrows():
            out_time = out_row['When']
            candidates = in_df[(in_df['When'] > out_time) & (~in_df['used']) & (in_df['date'] == out_time.date())]
            if not candidates.empty:
                in_row = candidates.iloc[0]
                in_time = in_row['When']
                durasi = (in_time - out_time).total_seconds() / 60
                if durasi <= 210:
                    results.append({
                        'Nama': name,
                        'OUT_When': out_time,
                        'IN_When': in_time,
                        'Durasi_menit': durasi,
                    })
                    in_df.loc[in_row.name, 'used'] = True

    df_pairing = pd.DataFrame(results)
    df_pairing = df_pairing.sort_values(by=['Nama', 'OUT_When']).reset_index(drop=True)

    # === DISPLAY RESULTS ===
    st.subheader("Hasil Pairing")
    st.dataframe(df_pairing)

    # === EXPORT TO GOOGLE SHEET ===
    if st.button("Save to Google Spreadsheet"):
        try:
            sheet = client.open_by_key("1byoXaizt1OLFJNDcWyf9DHNe0aoxfMmsId54o2L9lFE").worksheet("Sheet1")
            sheet.clear()
            sheet.update([df_pairing.columns.values.tolist()] + df_pairing.values.tolist())
            st.success("✅ Data berhasil disimpan ke Google Sheets")
        except Exception as e:
            st.error(f"❌ Gagal menyimpan ke spreadsheet: {e}")

    # === DOWNLOAD EXCEL ===
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_pairing.to_excel(writer, index=False)
    st.download_button("Download Excel", data=buffer.getvalue(), file_name="hasil_pairing.xlsx")

else:
    st.info("Silakan upload file untuk mulai analisis.")

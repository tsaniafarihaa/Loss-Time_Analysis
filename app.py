import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import time, datetime
import io

st.set_page_config(page_title="Loss Time Analysis", layout="wide")

st.title("Loss Time Analysis Dashboard")

st.markdown("""
Aplikasi ini menganalisis log akses untuk menghitung waktu yang dihabiskan di luar area yang ditentukan.
Upload file Excel dengan log akses yang berisi catatan masuk/keluar karyawan.
""")

# Upload file
uploaded_file = st.file_uploader("Upload file Excel log akses", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Fungsi pemrosesan utama
    @st.cache_data
    def process_data(file):
        try:
            df = pd.read_excel(file)
            st.success("File berhasil diupload!")
            
            # Tampilkan sampel data
            st.subheader("Sampel Data")
            st.dataframe(df.head())
            
            # Langkah 1: Ekstrak ID dari kolom 'Who'
            df['ID'] = df['Who'].str.extract(r'(\d+)', expand=False)
            
            # Langkah 2: Daftar ID yang dikecualikan (seperti di notebook)
            excluded_ids = [
                "100022883", "100028130", "100033411", "100034976", "100037620", "100038232", "100039869", "100040212",
                "100041020", "100041781", "100041693", "100041696", "100043374", "100043704", "100044690", "100047163",
                "100048938", "100049561", "100052278", "100054726", "100055362", "100057600", "100048754", "100062079",
                "100064653", "100064655", "2434673", "2588319", "2946384", "3013011", "2436060", "3227019", "3229196",
                "3242216", "3252881", "3251966", "3294315", "3301049", "3301547", "3334845", "3328895", "3442257",
                "3524553", "3568419", "3616448", "3616442", "3632147", "3641947", "3650221", "3667379", "3644405",
                "3677808", "3631443", "3706789", "3793827", "3808739", "3810788", "3812296", "3810787", "3883004",
                "3929368", "3988422", "4002219", "4039413", "4094250", "4117222", "3293057", "3645325", "3937753",
                "3957764", "3808729", "4033286", "4115553", "4116558", "4090423", "4090422", "100046032", "100046033",
                "2279149", "3329836", "3329835", "3329837", "3451556", "3510265", "3644227", "4048557", "4091122",
                "4036324"
            ]
            
            # Langkah 3: Filter baris dengan ID yang ada di daftar pengecualian
            df = df[~df['ID'].isin(excluded_ids)]
            
            # Hanya ambil baris dengan PEDESTRIAN IN/OUT
            df = df[df['Where'].str.contains('PEDESTRIAN IN|PEDESTRIAN OUT', na=False)]
            
            # Hapus entri dengan nama generik
            df = df[~df['Who'].str.contains('Temporary|Kartu|Costumer|999999', na=False, case=False)]
            
            # Hapus kolom yang tidak bernama jika ada
            if 'Unnamed: 5' in df.columns:
                df = df.drop(columns=['Unnamed: 5'])
            
            # Parse kolom datetime
            df['When'] = pd.to_datetime(df['When'], format='%d/%m/%Y %H:%M', errors='coerce')
            
            # Hapus baris dengan parsing tanggal yang gagal
            df = df.dropna(subset=['When'])
            
            # Tetapkan IN / OUT
            df['Direction'] = df['Where'].apply(lambda x: 'IN' if 'PEDESTRIAN IN' in x else 'OUT')
            
            # Hitung IN dan OUT per Who
            in_count = df[df['Direction'] == 'IN'].groupby('Who').size()
            out_count = df[df['Direction'] == 'OUT'].groupby('Who').size()
            
            # Gabungkan menjadi satu DataFrame
            direction_counts = pd.DataFrame({'IN': in_count, 'OUT': out_count}).fillna(0)
            
            # Simpan hanya yang IN = OUT
            valid_whos = direction_counts[direction_counts['IN'] == direction_counts['OUT']].index
            df = df[df['Who'].isin(valid_whos)]
            
            # Ekstrak nama
            df['Nama'] = df['Who'].str.extract(r',\s*(.*)$')[0]
            
            # Bersihkan spasi awal/akhir
            df['Nama'] = df['Nama'].str.strip()
            
            # Filter nama spesifik
            df = df[df['Nama'] != 'Dari A011']
            
            # Proses pemasangan
            pairings = []
            
            # Iterasi untuk setiap orang
            for name, group in df.groupby('Nama'):
                group = group.reset_index(drop=True)
                used_idx = set()
            
                for i, row in group.iterrows():
                    if row['Direction'] == 'OUT' and i not in used_idx:
                        out_time = row['When']
                        out_idx = i
            
                        # Cari IN pertama setelah OUT pada hari yang sama
                        for j in range(i+1, len(group)):
                            if group.loc[j, 'Direction'] == 'IN' and group.loc[j, 'When'].date() == out_time.date() and j not in used_idx:
                                in_time = group.loc[j, 'When']
                                duration = (in_time - out_time).total_seconds() / 60
            
                                pairings.append({
                                    'Nama': name,
                                    'Tanggal': out_time.date(),
                                    'OUT': out_time,
                                    'IN': in_time,
                                    'Durasi (menit)': round(duration)
                                })
            
                                used_idx.update([out_idx, j])
                                break  # berhenti setelah menemukan pasangan pertama
            
            # Buat DataFrame berpasangan
            paired_df = pd.DataFrame(pairings)
            
            # Filter: hapus durasi 0, lewati durasi > 210 menit
            paired_df = paired_df[(paired_df['Durasi (menit)'] > 0) & (paired_df['Durasi (menit)'] <= 210)]
            
            # Definisikan slot shift
            shift_slots = {
                'Shift 1': [
                    {'start': time(6,30), 'end': time(9,30), 'jenis': '-'},
                    {'start': time(9,30), 'end': time(11,0), 'jenis': 'Tea Break', 'durasi': 20},
                    {'start': time(11,0), 'end': time(13,30), 'jenis': 'Ishoma', 'durasi': 45},
                    {'start': time(13,30), 'end': time(14,0), 'jenis': '-'}
                ],
                'Shift 2': [
                    {'start': time(14,30), 'end': time(15,0), 'jenis': '-'},
                    {'start': time(15,0), 'end': time(16,40), 'jenis': 'Ashar', 'durasi': 20},
                    {'start': time(16,40), 'end': time(18,0), 'jenis': '-'},
                    {'start': time(18,0), 'end': time(19,0), 'jenis': 'Ishoma', 'durasi': 45},
                    {'start': time(19,0), 'end': time(20,30), 'jenis': '-'},
                    {'start': time(20,30), 'end': time(21,0), 'jenis': 'Isya', 'durasi': 20},
                    {'start': time(21,0), 'end': time(22,0), 'jenis': '-'}
                ],
                'Shift 3': [
                    {'start': time(0,0), 'end': time(1,30), 'jenis': 'Makan', 'durasi': 30},
                    {'start': time(1,30), 'end': time(4,0), 'jenis': '-'},
                    {'start': time(4,0), 'end': time(5,20), 'jenis': 'Subuh', 'durasi': 30},
                    {'start': time(5,20), 'end': time(6,0), 'jenis': '-'},
                    {'start': time(22,30), 'end': time(23,59,59), 'jenis': '-'}
                ]
            }
            
            # Tentukan shift berdasarkan waktu OUT
            def determine_shift(out_time):
                hour = out_time.hour
                if 6 <= hour < 14: return 'Shift 1'
                elif 14 <= hour < 22: return 'Shift 2'
                else: return 'Shift 3'
            
            # Hitung tumpang tindih menit antara dua rentang waktu
            def overlap_minutes(start1, end1, start2, end2):
                latest_start = max(start1, start2)
                earliest_end = min(end1, end2)
                overlap = (earliest_end - latest_start).total_seconds() / 60
                return max(0, overlap)
            
            # Hitung kategori dan waktu hilang
            def calculate_loss_time(row):
                shift = determine_shift(row['OUT'])
                slots = shift_slots[shift]
                total_istirahat = 0
                kategori = 'Kerja'
            
                for slot in slots:
                    slot_start = datetime.combine(row['OUT'].date(), slot['start'])
                    slot_end = datetime.combine(row['OUT'].date(), slot['end'])
                    keluar = row['OUT']
                    masuk = row['IN']
            
                    overlap = overlap_minutes(keluar, masuk, slot_start, slot_end)
                    if overlap > 0 and slot['jenis'] != '-':
                        total_istirahat += min(overlap, slot.get('durasi', 0))
                        kategori = slot['jenis']  # Ambil salah satu kategori yang cocok
            
                loss = row['Durasi (menit)'] - total_istirahat
                return pd.Series([kategori, max(0, loss)], index=['Kategori', 'Loss Time (menit)'])
            
            # Terapkan perhitungan
            paired_df[['Kategori', 'Loss Time (menit)']] = paired_df.apply(calculate_loss_time, axis=1)
            
            return paired_df
        
        except Exception as e:
            st.error(f"Error memproses file: {e}")
            return None
    
    # Proses data
    result_df = process_data(uploaded_file)
    
    if result_df is not None:
        # Visualisasi dan analisis
        st.subheader("Hasil Analisis Data")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write(f"Total catatan: {len(result_df)}")
            st.write(f"Total karyawan: {result_df['Nama'].nunique()}")
            st.write(f"Rentang tanggal: {result_df['Tanggal'].min()} sampai {result_df['Tanggal'].max()}")
        
        with col2:
            st.write(f"Rata-rata waktu hilang: {result_df['Loss Time (menit)'].mean():.2f} menit")
            st.write(f"Waktu hilang minimum: {result_df['Loss Time (menit)'].min():.2f} menit")
            st.write(f"Waktu hilang maksimum: {result_df['Loss Time (menit)'].max():.2f} menit")
        
        # Tampilkan data yang telah diproses
        st.subheader("Data yang Diproses")
        st.dataframe(result_df)
        
        # 10 karyawan teratas dengan waktu hilang tertinggi
        st.subheader("10 Karyawan dengan Waktu Hilang Tertinggi")
        
        top10 = result_df.groupby('Nama')['Loss Time (menit)'].sum().sort_values(ascending=False).head(10) / 60
        
        fig, ax = plt.subplots(figsize=(12, 6))
        top10.plot(kind='bar', color='orchid', ax=ax)
        plt.title('10 Karyawan dengan Waktu Hilang Tertinggi (Jam)')
        plt.xlabel('Nama')
        plt.ylabel('Waktu Hilang (Jam)')
        plt.xticks(rotation=45)
        plt.grid(axis='y', linestyle='--', alpha=0.6)
        plt.tight_layout()
        st.pyplot(fig)
        
        # Waktu hilang berdasarkan kategori
        st.subheader("Waktu Hilang Berdasarkan Kategori")
        
        category_loss = result_df.groupby('Kategori')['Loss Time (menit)'].sum().sort_values(ascending=False) / 60
        
        fig, ax = plt.subplots(figsize=(10, 6))
        category_loss.plot(kind='bar', colormap='viridis', ax=ax)
        plt.title('Total Waktu Hilang Berdasarkan Kategori (Jam)')
        plt.xlabel('Kategori')
        plt.ylabel('Waktu Hilang (Jam)')
        plt.xticks(rotation=45)
        plt.grid(axis='y', linestyle='--', alpha=0.6)
        plt.tight_layout()
        st.pyplot(fig)
        
        # Distribusi waktu
        st.subheader("Analisis Distribusi Waktu")
        
        fig, ax = plt.subplots(figsize=(12, 6))
        result_df['Hour'] = result_df['OUT'].dt.hour
        sns.histplot(data=result_df, x='Hour', bins=24, kde=True, ax=ax)
        plt.title('Distribusi Waktu Keluar (OUT) Berdasarkan Jam')
        plt.xlabel('Jam Hari')
        plt.ylabel('Jumlah')
        plt.xticks(range(0, 24))
        plt.grid(axis='y', linestyle='--', alpha=0.6)
        st.pyplot(fig)
        
        # Unduh data yang diproses
        csv = result_df.to_csv(index=False)
        st.download_button(
            label="Unduh data sebagai CSV",
            data=csv,
            file_name="analisis_waktu_hilang.csv",
            mime="text/csv",
        )
        
        # Unduh Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            result_df.to_excel(writer, sheet_name='Analisis Waktu Hilang', index=False)
        
        st.download_button(
            label="Unduh data sebagai Excel",
            data=buffer.getvalue(),
            file_name="analisis_waktu_hilang.xlsx",
            mime="application/vnd.ms-excel",
        )

else:
    st.info("Silakan unggah file Excel untuk memulai analisis.")

# Tambahkan informasi tentang aplikasi
st.sidebar.header("Tentang")
st.sidebar.info("""
Aplikasi ini menganalisis log akses untuk menghitung waktu hilang.
Aplikasi memproses catatan masuk/keluar untuk mengidentifikasi kapan karyawan meninggalkan area yang ditentukan.
Aplikasi dapat mengidentifikasi waktu istirahat, waktu sholat, dan ketidakhadiran yang tidak sah.
""")

st.sidebar.header("Cara Penggunaan")
st.sidebar.info("""
1. Unggah file Excel yang berisi log akses
2. Tinjau data yang telah diproses
3. Analisis visualisasi
4. Unduh hasil untuk analisis lebih lanjut
""")

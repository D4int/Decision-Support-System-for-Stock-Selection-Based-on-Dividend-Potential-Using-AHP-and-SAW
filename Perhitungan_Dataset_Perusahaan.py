import streamlit as st
import pandas as pd
import numpy as np

# Path ke file dataset
DATASET_PATH = "C:/Users/algha/Downloads/Project SPK/Project SPK/Dataset Original.xlsx"

# Fungsi untuk membaca dataset
@st.cache
def load_data():
    excel_data = pd.ExcelFile(DATASET_PATH)
    sheets = {sheet_name: excel_data.parse(sheet_name) for sheet_name in excel_data.sheet_names}
    return sheets

# Fungsi normalisasi matriks untuk AHP
def normalize_matrix(matrix):
    max_values = matrix.max(axis=0)  # Ambil nilai maksimum dari setiap kolom
    normalized_matrix = matrix / max_values  # Normalisasi dengan membagi setiap elemen dengan nilai maksimum kolom
    return normalized_matrix

# Fungsi perhitungan AHP
def ahp_process(matrix):
    # Hitung jumlah kolom (menurun)
    column_totals = matrix.sum(axis=0)
    
    # Normalisasi matriks (priority matrix)
    priority_matrix = matrix / column_totals
    
    # Hitung jumlah baris (ke kanan)
    row_totals = priority_matrix.sum(axis=1)
    
    # Hitung rata-rata baris untuk bobot kriteria
    row_means = priority_matrix.mean(axis=1)
    
    return column_totals, priority_matrix, row_totals, row_means

# Fungsi untuk SAW
def saw_ranking(data, weights):
    normalized_data = data / data.max(axis=0)  # Normalisasi data
    scores = np.dot(normalized_data, weights)  # Kalkulasi skor dengan metode SAW
    return scores

# Fungsi untuk membersihkan data dari tipe yang tidak numerik
def clean_data(data):
    numeric_data = data.select_dtypes(include=[np.number])  # Hanya pilih kolom numerik
    return numeric_data

# Fungsi untuk menghitung jumlah total vertikal
def calculate_vertical_totals(matrix):
    return matrix.sum(axis=0)  # Hitung jumlah setiap kolom

# Inisialisasi matriks perbandingan berpasangan
if 'input_matrix' not in st.session_state:
    st.session_state.input_matrix = np.array([
        [1, 3, 5],        # DPS terhadap DPS, DY, ROE
        [1/3, 1, 3],      # DY terhadap DPS, DY, ROE
        [1/5, 1/3, 1]     # ROE terhadap DPS, DY, ROE
    ])

# Fungsi utama aplikasi
st.title("Aplikasi SPK: AHP dan SAW dengan Dataset Langsung")
sheets = load_data()
sheet_names = list(sheets.keys())

# Pilih sheet
selected_sheet = st.selectbox("Pilih Dataset", sheet_names)
data = sheets[selected_sheet]

st.write(f"Dataset: {selected_sheet}")
st.dataframe(data)

# Bersihkan data dari kolom non-numerik
cleaned_data = clean_data(data)

# Proses AHP
if st.button("Proses AHP - Normalisasi Data"):
    st.subheader("Matriks Normalisasi AHP")
    try:
        matrix = cleaned_data.values  # Ambil data numerik
        normalized_matrix = normalize_matrix(matrix)
        st.session_state.normalized_matrix = normalized_matrix  # Simpan normalized_matrix di session_state
        st.dataframe(pd.DataFrame(normalized_matrix, columns=cleaned_data.columns, index=data.index))
    except Exception as e:
        st.error(f"Error: {e}")

# Tombol untuk memproses AHP
if st.button("Pembobotan AHP"):
    try:
        matrix = st.session_state.input_matrix

        # Langkah 1: Matriks Perbandingan Berpasangan Awal
        st.subheader("Langkah 1: Matriks Perbandingan Berpasangan (Awal)")
        column_totals = calculate_vertical_totals(matrix)

        # Proses perhitungan AHP
        column_totals, priority_matrix, row_totals, weights = ahp_process(matrix)
        st.session_state.weights = weights  # Simpan hasil bobot ke session_state

        # Gabungkan hasil dalam tabel
        combined_table = pd.DataFrame(priority_matrix, columns=["DPS", "DY", "ROE"], index=["DPS", "DY", "ROE"])
        combined_table["Jumlah"] = row_totals
        combined_table["Prioritas"] = weights  # Tambahkan kolom Prioritas
        
        # Tampilkan tabel gabungan
        st.subheader("Langkah 2-5: Tabel Normalisasi dan Bobot Kriteria")
        st.dataframe(combined_table)

        # Hitung Consistency Index dan Consistency Ratio
        lambda_max = (column_totals * weights).sum()
        consistency_index = (lambda_max - 3) / (3 - 1)
        consistency_ratio = consistency_index / 0.58  # 0.58 adalah nilai indeks konsistensi untuk 3 kriteria

        st.write("Consistency Index:", consistency_index)
        st.write("Consistency Ratio:", consistency_ratio)
    except Exception as e:
        st.error(f"Error: {e}")

# Tabel untuk Perankingan SAW
if st.button("Perankingan SAW"):
    try:
        if "normalized_matrix" in st.session_state and "weights" in st.session_state:
            normalized_matrix = st.session_state["normalized_matrix"]
            weights = st.session_state["weights"]

            # Hitung hasil kali normalisasi dengan bobot
            hasil_kali = normalized_matrix * weights
            scores = hasil_kali.sum(axis=1)  # Jumlahkan hasil kali per baris

            # Membuat tabel hasil
            result_table = pd.DataFrame({
                "Tanggal Ex-Dividen": data["Tanggal Ex-Dividen"] if "Tanggal Ex-Dividen" in data.columns else ["Tidak tersedia"] * len(scores),
                "Skor Total": scores
            })
            st.subheader("Hasil Perankingan SAW")
            st.dataframe(result_table)

        else:
            st.error("Silakan lakukan normalisasi data terlebih dahulu.")
    except Exception as e:
        st.error(f"Error: {e}")

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

# Fungsi normalisasi AHP
def normalize_matrix(matrix):
    max_values = matrix.max(axis=0)  # Ambil nilai maksimum dari setiap kolom
    normalized_matrix = matrix / max_values  # Bagi setiap elemen dengan nilai maksimum kolom
    return normalized_matrix

# Fungsi perhitungan AHP
def ahp_process(matrix):
    # Langkah 2: Hitung jumlah kolom (menurun)
    column_totals = matrix.sum(axis=0)
    
    # Langkah 3: Normalisasi matriks (priority matrix)
    priority_matrix = matrix / column_totals
    
    # Langkah 4: Hitung jumlah baris (ke kanan)
    row_totals = priority_matrix.sum(axis=1)
    
    # Langkah 5: Hitung rata-rata baris untuk bobot kriteria
    row_means = priority_matrix.mean(axis=1)
    
    return column_totals, priority_matrix, row_totals, row_means

def calculate_vertical_totals(matrix):
    return matrix.sum(axis=0)  # Menjumlahkan setiap kolom

# Inisialisasi matriks perbandingan berpasangan
if 'input_matrix' not in st.session_state:
    st.session_state.input_matrix = np.array([
        [1, 3, 5],        # DPS terhadap DPS, DY, ROE
        [1/3, 1, 3],      # DY terhadap DPS, DY, ROE
        [1/5, 1/3, 1]     # ROE terhadap DPS, DY, ROE
    ])

# Fungsi SAW untuk perankingan
def saw_ranking(data, weights):
    normalized_data = data / data.max(axis=0)  # Normalisasi untuk SAW
    scores = np.dot(normalized_data, weights)
    return scores

# Fungsi untuk membersihkan data dari tipe yang tidak numerik
def clean_data(data):
    numeric_data = data.select_dtypes(include=[np.number])  # Ambil hanya kolom numerik
    return numeric_data

# Load dataset
st.title("Aplikasi AHP dan SAW dengan Dataset Langsung")
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
        matrix = cleaned_data.values  # Ambil hanya data numerik
        normalized_matrix = normalize_matrix(matrix)
        st.session_state.normalized_matrix = normalized_matrix  # Simpan normalized_matrix di session_state
        st.dataframe(pd.DataFrame(normalized_matrix, columns=cleaned_data.columns, index=data.index))
    except Exception as e:
        st.error(f"Error: {e}")

# Tombol untuk memproses AHP
if st.button("Pembobotan AHP"):
    try:
        matrix = st.session_state.input_matrix
        
        # Langkah 1: Tampilkan matriks perbandingan berpasangan awal
        st.subheader("Langkah 1: Matriks Perbandingan Berpasangan (Awal)")

        # Hitung jumlah kolom (menurun) untuk ditambahkan sebagai baris baru
        column_totals = matrix.sum(axis=0)

        # Hitung total vertikal untuk setiap kolom
        vertical_totals = calculate_vertical_totals(matrix)

        # Tambahkan baris baru "Jumlah" ke matriks
        matrix_with_totals = np.vstack([matrix, column_totals])
        row_labels = ["DPS", "DY", "ROE", "Jumlah"]  # Tambahkan label untuk baris baru
        
        # Tampilkan tabel dengan baris "Jumlah"
        st.dataframe(pd.DataFrame(matrix_with_totals, 
                                  columns=["DPS", "DY", "ROE"], 
                                  index=row_labels))
        
        # Proses perhitungan AHP
        column_totals, priority_matrix, row_totals, weights = ahp_process(matrix)

        # Simpan hasil bobot ke session_state
        st.session_state.weights = weights

        # Gabungkan hasil dari langkah 2, 3, 4, dan 5 dalam satu tabel
        combined_table = pd.DataFrame(priority_matrix, columns=["DPS", "DY", "ROE"], index=["DPS", "DY", "ROE"])
        combined_table["Jumlah"] = row_totals
        combined_table["Prioritas"] = weights  # Tambahkan kolom Prioritas
        
        # Tambahkan baris jumlah total vertikal
        column_totals_row = pd.DataFrame([vertical_totals.tolist() + ["", ""]], 
                                         columns=["DPS", "DY", "ROE", "Jumlah", "Prioritas"], 
                                         index=["Jumlah"])
        
        final_table = pd.concat([combined_table, column_totals_row])
        
        # Tampilkan tabel gabungan
        st.subheader("Langkah 2 : Bagi setiap elemen dalam matriks dengan total kolomnya masing-masing")
        st.dataframe(final_table)
        
        # Membuat tabel Langkah 3
        st.subheader("Langkah 3 : Mencari λ maks untuk menghitung Consistency Ratio")
        
        # Menghitung tabel Langkah 3
        langkah3_data = np.zeros((3, 3))  # Inisialisasi matriks dengan nol
        for i in range(3):  # Mengiterasi setiap baris (DPS, DY, ROE)
            for j in range(3):  # Mengiterasi setiap kolom (DPS, DY, ROE)
                langkah3_data[i, j] = matrix[i, j] * weights[j]  # Pembagian elemen dengan prioritas kolom
                
        # Menjumlahkan setiap baris untuk hasil perkalian
        hasil_perkalian = langkah3_data.sum(axis=1)
        
        # Menghitung Hasil
        hasil = hasil_perkalian / weights  # Pembagian hasil_perkalian dengan prioritas
        
        # Membuat DataFrame untuk tabel Langkah 3
        langkah3_df = pd.DataFrame(langkah3_data, columns=["DPS", "DY", "ROE"], index=["DPS", "DY", "ROE"])
        langkah3_df["Hasil Perkalian"] = hasil_perkalian
        langkah3_df["Prioritas"] = weights  # Menambahkan kolom Prioritas
        langkah3_df["Hasil"] = hasil  # Menambahkan kolom Hasil
        
        # Menghitung jumlah untuk kolom Hasil
        langkah3_df.loc["Jumlah"] = ["", "", "", "", "", hasil.sum()]  # Hanya hasil yang dijumlahkan
        
        # Menambahkan baris Kriteria dan λ maks
        langkah3_df.loc["Kriteria"] = ["", "", "", "", "", 3]  # Nilai Hasil Kriteria diisi 3
        lambda_max = hasil.sum() / 3  # λ maks dihitung dari jumlah Hasil / 3
        langkah3_df.loc["λ maks"] = ["", "", "", "", "", lambda_max]  # Menambahkan λ maks ke tabel
        
        # Tampilkan tabel Langkah 3
        st.dataframe(langkah3_df)

        # Langkah 4: Hitung Consistency Index dan Consistency Ratio
        consistency_index = (lambda_max - 3) / (3 - 1)
        consistency_ratio = consistency_index / 0.58  # 0.58 adalah nilai indeks konsistensi untuk 3 kriteria

        # Buat DataFrame untuk Consistency Index dan Consistency Ratio
        consistency_table = pd.DataFrame({
            "": ["Consistency Index", "Consistency Ratio"],
            "Nilai": [consistency_index, consistency_ratio]
        })

        # Tampilkan tabel Consistency
        st.subheader("Langkah 4: Consistency Index dan Consistency Ratio")
        st.dataframe(consistency_table)

        # Langkah 5: Tabel Bobot
        weight_table = pd.DataFrame({
            "Kriteria": ["DPS", "DY (Dividend Yield)", "ROE (Return on Equity)"],
            "Bobot": weights
            })
        # Tampilkan tabel bobot
        st.subheader("Langkah 5: Tabel Bobot AHP")
        st.dataframe(weight_table)

    except Exception as e:
        st.error(f"Error: {e}")

# Tabel untuk Perankingan SAW
if st.button("Perankingan SAW"):
    try:
        if "normalized_matrix" in st.session_state and "weights" in st.session_state:
            normalized_matrix = st.session_state["normalized_matrix"]
            weights = st.session_state["weights"]

            # Hitung hasil kali normalisasi dengan bobot
            hasil_kali = normalized_matrix * weights  # Elemen matriks dikali dengan bobot
            
            # Membuat DataFrame hasil normalisasi dan hasil kali
            hasil_dataframe = pd.DataFrame({
                "Normalisasi (DPS)": normalized_matrix[:, 0],
                "Normalisasi (DY)": normalized_matrix[:, 1],
                "Normalisasi (ROE)": normalized_matrix[:, 2],
                "Hasil Kali (DPS)": hasil_kali[:, 0],
                "Hasil Kali (DY)": hasil_kali[:, 1],
                "Hasil Kali (ROE)": hasil_kali[:, 2],
            }, index=data.index)  # Gunakan indeks asli dataset

            # Tampilkan tabel pertama
            st.subheader("Hasil Perankingan SAW")
            st.dataframe(hasil_dataframe)

            # Hitung Skor Total
            scores = hasil_kali.sum(axis=1)  # Jumlahkan hasil kali per baris

            # Ambil kolom Tanggal Ex-Dividen dari dataset original
            if "Tanggal Ex-Dividen" in data.columns:
                tanggal_ex_dividen = data["Tanggal Ex-Dividen"]
            else:
                st.error("Kolom 'Tanggal Ex-Dividen' tidak ditemukan di dataset.")
                tanggal_ex_dividen = ["Tidak tersedia"] * len(scores)  # Placeholder jika kolom tidak ada

            # Membuat tabel baru untuk skor total
            result_table = pd.DataFrame({
                "Tanggal Ex-Dividen": tanggal_ex_dividen,
                "Skor Total": scores
            })

            # Tampilkan tabel kedua
            st.subheader("Menjumlahkan hasil perkalian untuk mendapatkan skor total")
            st.dataframe(result_table)

            # Hitung Jumlah Total dari Skor Total
            jumlah_total = scores.sum()  # Hitung total skor
            jumlah_table = pd.DataFrame({
                "Data Perusahaan": ["Jumlah"],
                "Skor Total": [jumlah_total]
            })

            # Tampilkan tabel jumlah
            st.subheader("Data Perusahaan")
            st.dataframe(jumlah_table)

        else:
            st.error("Silakan lakukan normalisasi data terlebih dahulu.")
    except Exception as e:
        st.error(f"Error: {e}")

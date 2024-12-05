import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Baca data dari Excel
file_path = "C:/Users/algha/Downloads/Project SPK/Project SPK/Visualisasikan.xlsx"  # Ganti dengan path file Anda
data = pd.read_excel(file_path, sheet_name='Perankingan')

# Judul aplikasi
st.title("Ranking Perusahaan Berdasarkan Kriteria")

# Menampilkan data
st.write("Data Ranking:")
st.dataframe(data)

# Pilih kolom untuk grafik
x_column = st.selectbox("Pilih kolom untuk sumbu X:", data.columns)
y_column = st.selectbox("Pilih kolom untuk sumbu Y:", data.columns)

# Buat grafik
plt.figure(figsize=(10, 5))

# Periksa tipe data untuk X dan Y
if pd.api.types.is_numeric_dtype(data[x_column]) and pd.api.types.is_numeric_dtype(data[y_column]):
    # Jika kedua kolom adalah numerik, gunakan scatter plot
    plt.scatter(data[x_column], data[y_column], color='blue', alpha=0.7)
    plt.xlabel(x_column)
    plt.ylabel(y_column)
    plt.title(f'Grafik {y_column} vs {x_column}')
elif pd.api.types.is_object_dtype(data[x_column]) and pd.api.types.is_numeric_dtype(data[y_column]):
    # Jika X adalah kategori dan Y adalah numerik, gunakan bar chart horizontal
    sorted_data = data.sort_values(by=y_column, ascending=False)
    plt.barh(sorted_data[x_column], sorted_data[y_column], color='skyblue')
    plt.xlabel(y_column)
    plt.ylabel(x_column)
    plt.title(f'Grafik {y_column} vs {x_column}')
elif pd.api.types.is_numeric_dtype(data[x_column]) and pd.api.types.is_object_dtype(data[y_column]):
    # Jika X adalah numerik dan Y adalah kategori, gunakan bar chart vertikal
    sorted_data = data.sort_values(by=x_column, ascending=True)
    plt.bar(sorted_data[x_column], sorted_data[y_column], color='orange')
    plt.xlabel(x_column)
    plt.ylabel(y_column)
    plt.title(f'Grafik {y_column} vs {x_column}')
else:
    # Jika kedua kolom adalah kategori, beri peringatan
    st.warning("Tidak dapat membuat grafik dengan dua kolom bertipe kategori.")

# Tampilkan grafik
st.pyplot(plt)

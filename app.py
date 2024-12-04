from flask import Flask, request, render_template, redirect, url_for, send_file, jsonify
import pandas as pd
import numpy as np
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

@app.route("/")
def index():
    # Halaman awal
    return render_template("index.html", result="", file_url="#")

@app.route('/download_result')
def download_result():
    try:
        result_filepath = os.path.join(RESULT_FOLDER, 'Hasil_Seleksi_Konsentrasi.xlsx')
        if os.path.exists(result_filepath):
            return send_file(result_filepath, as_attachment=True, download_name='Hasil_Seleksi_Konsentrasi.xlsx')
        else:
            return "File tidak ditemukan. Silakan ulangi proses terlebih dahulu.", 404
    except Exception as e:
        return f"Terjadi kesalahan saat mengunduh file: {e}", 500

@app.route('/process', methods=['POST'])
def process():
    try:
        # File upload
        file = request.files['file']
        kuota_ai = int(request.form['kuota_ai'])
        
        konsentrasi = int(request.form['konsentrasi'])

        # Menangkap prioritas dari form HTML
        prioritas_matkul = [
            (request.form.get('nama_matkul_1'), int(request.form.get('prioritas_1'))),
            (request.form.get('nama_matkul_2'), int(request.form.get('prioritas_2'))),
            (request.form.get('nama_matkul_3'), int(request.form.get('prioritas_3'))),
            (request.form.get('nama_matkul_4'), int(request.form.get('prioritas_4'))),
            (request.form.get('nama_matkul_5'), int(request.form.get('prioritas_5'))),
            (request.form.get('nama_matkul_6'), int(request.form.get('prioritas_6'))),
            (request.form.get('nama_matkul_7'), int(request.form.get('prioritas_7'))),
            (request.form.get('nama_matkul_8'), int(request.form.get('prioritas_8'))),
            (request.form.get('nama_matkul_9'), int(request.form.get('prioritas_9'))),]
        
        # Urutkan berdasarkan prioritas (elemen kedua dalam tuple)
        prioritas_matkul.sort(key=lambda x: x[1])

        # Ambil hanya nama mata kuliah setelah diurutkan
        prioritas_matkul = [nama for nama, prioritas in prioritas_matkul]
        
        print(f"Prioritas Nilai Matkul: {prioritas_matkul}")
        
        prioritas_pendukung = [
            (request.form.get('nama_pendukung_1'), int(request.form.get('prioritaspendukung_1'))),
            (request.form.get('nama_pendukung_2'), int(request.form.get('prioritaspendukung_2'))),
            (request.form.get('nama_pendukung_3'), int(request.form.get('prioritaspendukung_3'))),]
        
        # Urutkan berdasarkan prioritas (elemen kedua dalam tuple)
        prioritas_pendukung.sort(key=lambda x: x[1])

        # Ambil hanya nama mata kuliah setelah diurutkan
        prioritas_pendukung = [nama for nama, prioritas in prioritas_pendukung]
        
        print(f"Prioritas Nilai Pendukung: {prioritas_pendukung}")

        # Simpan file ke folder
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        # Load dataset
        data = pd.read_excel(filepath)

        # Proses data (tetap menggunakan logika yang Anda miliki di atas)
        data = data[['Nama Lengkap', 'Karya', 'SMK/SMA/MA', 'Hobi', 'Nilai Mata Kuliah Struktur Data', 
                     'Nilai Mata Kuliah Algoritma dan pemrograman dasar', 'Nilai Mata Kuliah Pemrograman Lanjut', 
                     'Nilai Mata Kuliah Statistik dan Probabilitas', 'Nilai Mata Kuliah Keamanan Informasi', 
                     'Nilai Mata Kuliah Jaringan Komputer', 'Nilai Mata Kuliah Arsitektur dan Organisasi Komputer', 
                     'Nilai Mata Kuliah Sistem Digital', 'Nilai Mata Kuliah Rangkaian Elektronika', 
                     'Pilihan Konsentrasi 1']]

        # Membuat dictionary untuk konversi nilai
        konversi_nilai = {
            'A': 10, 'A-': 9, 'B+': 8, 'B': 7, 'B-': 6, 'C+': 5, 'C': 4, 'C- hingga E': 2, 'Belum diprogramkan': 1
        }

        # Fungsi untuk mengkonversi nilai mata kuliah, Hobi, Karya, SMK/SMA/MA, dan Pilihan Konsentrasi
        def konversi(row):
            # Kolom nilai mata kuliah berada pada indeks 4 hingga 12 (kolom 10-18 adalah indeks ke 4 sampai 12)
            for kolom in data.columns[4:13]:  # Menyesuaikan dengan kolom 10 sampai 18
                if row[kolom] in konversi_nilai:
                    row[kolom] = konversi_nilai[row[kolom]]
                    
            hobi_keywords = [
            'Pemrograman',
            'Menganalisis data'
            ]
            
            # Hitung jumlah kata kunci yang ada di kolom 'Karya'
            matches = sum(1 for keyword in hobi_keywords if keyword in str(row['Hobi']))

            # Tentukan nilai berdasarkan jumlah kecocokan
            if matches == 2:
                row['Hobi'] = 3
            elif matches == 1:
                row['Hobi'] = 2
            else:
                row['Hobi'] = 1  # Jika tidak ada yang cocok
                        
            karya_keywords = [
            'Pernah membuat aplikasi berbasis Artificial Intelligence',
            'Pernah membuat Game',
            'aplikasi mobile',
            'Pernah membuat website',
            'Pernah membuat Sistem Informasi dengan tampilan menarik dan berfungsi dengan baik'
            ]

            # Hitung jumlah kata kunci yang ada di kolom 'Karya'
            matches = sum(1 for keyword in karya_keywords if keyword in str(row['Karya']))

            # Tentukan nilai berdasarkan jumlah kecocokan
            if matches == 4 or matches == 5:
                row['Karya'] = 9
            elif matches == 3:
                row['Karya'] = 7
            elif matches == 2:
                row['Karya'] = 5
            elif matches == 1:
                row['Karya'] = 3
            else:
                row['Karya'] = 1  # Jika tidak ada yang cocok

                        
            # Mengonversi nilai kolom SMK/SMA/MA
            if 'SMK' in str(row['SMK/SMA/MA']):
                row['SMK/SMA/MA'] = 5
            else:
                row['SMK/SMA/MA'] = 3
                        
            # Mengonversi nilai kolom Pilihan Konsentrasi 1
            if 'Artificial Intelligence' in str(row['Pilihan Konsentrasi 1']):
                row['Pilihan Konsentrasi 1'] = 1
            elif 'Network & Security' in str(row['Pilihan Konsentrasi 1']):
                row['Pilihan Konsentrasi 1'] = 2
            elif 'Embedded System' in str(row['Pilihan Konsentrasi 1']):
                row['Pilihan Konsentrasi 1'] = 3
            else:
                row['Pilihan Konsentrasi 1'] = 0  # Nilai default jika tidak cocok

            return row

        # Terapkan konversi nilai ke seluruh data
        data = data.apply(konversi, axis=1)

        # Menyaring data untuk hanya yang memiliki nilai "1" pada kolom "Pilihan Konsentrasi 1"
        data = data[data['Pilihan Konsentrasi 1'] == konsentrasi]
        
        if kuota_ai >= len(data):
            # Semua data akan diambil karena kuota lebih besar atau sama dengan jumlah data
            data_filtered = data

            # Simpan hasil ke file Excel
            result_filepath = os.path.join(RESULT_FOLDER, 'Hasil_Seleksi_Konsentrasi.xlsx')
            data_filtered.to_excel(result_filepath, index=True)

            # Konversi data ke tabel HTML
            result_html = data_filtered[['Nama Lengkap', 'Pilihan Konsentrasi 1']].to_html(
                classes="table table-striped table-bordered", index=True
            )

            # Kembalikan halaman dengan tabel HTML dan link unduhan
            return render_template(
                "index.html",
                result=result_html,
                file_url=url_for("download_result")
            )

        # Matriks perbandingan berpasangan untuk sub-kriteria Nilai Mata Kuliah
        mata_kuliah_matrix = np.array([[1, 3, 5, 7, 9, 9, 9, 9, 9], 
                                    [1/3, 1, 3, 5, 7, 7, 7, 7, 7],
                                    [1/5, 1/3, 1, 3, 5, 5, 5, 5, 5],
                                    [1/7, 1/5, 1/3, 1, 3, 3, 3, 3, 3],
                                    [1/9, 1/7, 1/5, 1/3, 1, 3, 3, 3, 3],
                                    [1/9, 1/7, 1/5, 1/3, 1/3, 1, 3, 3, 3],
                                    [1/9, 1/7, 1/5, 1/3, 1/3, 1/3, 1, 3, 3],
                                    [1/9, 1/7, 1/5, 1/3, 1/3, 1/3, 1/3, 1, 3],
                                    [1/9, 1/7, 1/5, 1/3, 1/3, 1/3, 1/3, 1/3, 1]])

        # Matriks perbandingan berpasangan untuk sub-kriteria Nilai Pendukung
        pendukung_matrix = np.array([[1, 3, 5], 
                                    [1/3, 1, 3], 
                                    [1/5, 1/3, 1]])

        # Fungsi untuk normalisasi matriks perbandingan
        def normalize_matrix(matrix):
            column_sum = matrix.sum(axis=0)
            normalized_matrix = matrix / column_sum
            return normalized_matrix

        # Normalisasi matriks dan hitung bobot
        def calculate_weights(matrix):
            normalized_matrix = normalize_matrix(matrix)
            weights = normalized_matrix.mean(axis=1)  # Rata-rata tiap baris sebagai bobot
            return weights

        # Menghitung bobot sub-kriteria untuk Nilai Mata Kuliah
        bobot_mata_kuliah = calculate_weights(mata_kuliah_matrix)
        print(f"Bobot Sub-Kriteria Nilai Mata Kuliah: {bobot_mata_kuliah}")

        # Menghitung bobot sub-kriteria untuk Nilai Pendukung
        bobot_pendukung = calculate_weights(pendukung_matrix)
        print(f"Bobot Sub-Kriteria Nilai Pendukung: {bobot_pendukung}")
        
        print(f"Data Kolom: {data.columns}")

        
        data_matkul = data[prioritas_matkul]
        
        print(f"Test: {data_matkul}")

        # Hitung skor akhir berdasarkan bobot dan nilai masing-masing kolom
        nilai_mata_kuliah_skore = (data_matkul.iloc[:, 0] ** bobot_mata_kuliah[0]) * \
                                    (data_matkul.iloc[:, 1] ** bobot_mata_kuliah[1]) * \
                                    (data_matkul.iloc[:, 2] ** bobot_mata_kuliah[2]) * \
                                    (data_matkul.iloc[:, 3] ** bobot_mata_kuliah[3]) * \
                                    (data_matkul.iloc[:, 4] ** bobot_mata_kuliah[4]) * \
                                    (data_matkul.iloc[:, 5] ** bobot_mata_kuliah[5]) * \
                                    (data_matkul.iloc[:, 6] ** bobot_mata_kuliah[6]) * \
                                    (data_matkul.iloc[:, 7] ** bobot_mata_kuliah[7]) * \
                                    (data_matkul.iloc[:, 8] ** bobot_mata_kuliah[8])

        print(f"Skor Akhir Mata Kuliah: {nilai_mata_kuliah_skore}")
        
        data_pendukung = data[prioritas_pendukung]
        
        print(f"Data Setelah Normalisasi: {data_pendukung}")
                                    
        skor_pendukung = (data_pendukung.iloc[:, 0] ** bobot_pendukung[0]) * \
                            (data_pendukung.iloc[:, 1] ** bobot_pendukung[1]) * \
                            (data_pendukung.iloc[:, 2] ** bobot_pendukung[2])
                            
        print(f"Normalisasi Pendukung: {skor_pendukung}")
        
        nilai = nilai_mata_kuliah_skore + skor_pendukung
        
        sum_nilai = np.sum(nilai)
        
        print(sum_nilai)
        
        hasil = nilai/sum_nilai
        
        print(f"Hasil: {hasil}")

        # Tambahkan total skor ke dalam DataFrame
        data['Total Skor'] = hasil

        # Mengurutkan data berdasarkan total skor
        data_sorted = data.sort_values(by='Total Skor', ascending=False)
        
        data_filtered = data_sorted.head(kuota_ai)

        # Simpan hasil ke file Excel
        result_filepath = os.path.join(RESULT_FOLDER, 'Hasil_Seleksi_Konsentrasi.xlsx')
        data_filtered.to_excel(result_filepath, index=True)

        # Konversi data ke tabel HTML
        result_html = data_filtered[['Nama Lengkap', 'Total Skor', 'Pilihan Konsentrasi 1']].to_html(classes="table table-striped table-bordered", index=True)
        
        result_file_path = "path/to/result.xlsx"

        print(result_html)  # Cek isi variabel result
        return render_template(
        "index.html",
        result=result_html,
        file_url=url_for("download_result") if result_file_path else "#"
    )

    except Exception as e:
        return render_template('index.html', result=f"Terjadi kesalahan: {e}")

if __name__ == '__main__':
    app.run(debug=True)

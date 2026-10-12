# 📄 Aplikasi Dokumentasi Kegiatan (Word Siap Cetak)

Aplikasi berbasis web untuk menyusun foto dokumentasi kegiatan secara otomatis ke dalam dokumen **Microsoft Word (.docx)** ukuran **A4 Portrait** dengan tata letak **Grid 3×3 Siap Cetak**, serta fitur pengiriman otomatis ke email.

---

## 🚀 Fitur Utama
1. **Tata Letak 3×3 Otomatis**: Menyusun 9 foto per halaman dengan rasio aspek asli yang terjaga (tidak gepeng/terdistorsi).
2. **Pembersihan Otomatis**:
   - Memfilter dan mengabaikan foto duplikat (hash detection).
   - Merapikan bilah hitam screenshot secara otomatis.
3. **Format Resmi Siap Cetak**:
   - Header terstruktur: `DOKUMENTASI KEGIATAN`, `Nama Kegiatan`, `Tanggal`.
   - Ukuran kertas standar A4 Portrait dengan margin cetak yang presisi.
4. **Unduh Langsung & Kirim Email**:
   - File `.docx` dapat langsung diunduh dari browser.
   - Pilihan kirim otomatis sebagai lampiran email ke alamat tujuan (default: `hitamair4@gmail.com`).

---

## 💻 Cara Menjalankan di Komputer / Laptop

### Pengguna Windows:
1. Pastikan Python sudah terinstal di komputer ([python.org](https://www.python.org)).
2. Ekstrak file zip ini ke folder baru.
3. Klik 2x pada file **`run_app.bat`**.
4. Halaman aplikasi akan otomatis terbuka di browser Anda (`http://localhost:8501`).

### Pengguna Mac / Linux:
1. Buka Terminal di dalam folder ini.
2. Jalankan perintah:
   ```bash
   pip install -r requirements.txt
   streamlit run app.py
   ```
3. Browser akan otomatis terbuka.

---

## 🌐 Cara Deploy Gratis ke Internet (Agar Bisa Diakses Siapa Saja)
1. Buat akun di [GitHub](https://github.com) dan unggah folder ini ke repositori baru.
2. Buka [share.streamlit.io](https://share.streamlit.io) (gratis dari Streamlit).
3. Sambungkan repositori GitHub Anda dan klik **Deploy**.
4. Anda akan mendapatkan tautan web publik (misal: `https://dokumentasi-kegiatan.streamlit.app`) yang bisa diakses dari HP atau laptop kapan saja tanpa menyalakan komputer sendiri.

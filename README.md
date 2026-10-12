# 📸 Dokumentasi Foto — Simpan Momen, Abadikan Cerita (Versi 2.0)

Aplikasi web modern berbasis **Python & Streamlit** untuk mengotomatiskan pembuatan dokumen foto kegiatan resmi siap cetak (**Microsoft Word .docx** & **PDF**) ukuran A4 Portrait dengan tata letak Grid 3×3 serta integrasi penyimpanan terstruktur ke **Google Drive**.

Dirancang khusus untuk kebutuhan perangkat desa, panitia kegiatan, sekolah, instansi, dan masyarakat.

---

## 🌟 Fitur Utama & Penyempurnaan Baru

1. **Dashboard & Statistik Real-Time**:
   - Menampilkan total kegiatan, total dokumen, foto terproses, dan kapasitas Google Drive.
   - Daftar riwayat kegiatan terbaru secara langsung dari database lokal persisten (SQLite).
2. **Alur Kerja Pembuatan Dokumen (Wizard 4 Langkah)**:
   - **Langkah 1: Informasi**: Input nama kegiatan, tanggal (date picker), lokasi/keterangan, format dokumen, dan tombol simpan Drive.
   - **Langkah 2: Pengelolaan Foto**: Drag & drop multi-foto, deteksi foto duplikat otomatis (MD5 hash), pembersihan bilah hitam tangkapan layar, dan opsi hapus foto.
   - **Langkah 3: Pratinjau Dokumen**: Visual A4 mockup interaktif dengan kontrol pemilihan jumlah kolom (2, 3, 4, 6 kolom), orientasi foto (Proporsional vs Isi Bingkai), dan estimasi jumlah halaman cetak.
   - **Langkah 4: Hasil & Penyimpanan**: Tombol unduh langsung DOCX & PDF, serta informasi folder Google Drive.
3. **Penyimpanan Terstruktur Google Drive**:
   - Otomatis membuat hierarki folder rapi:
     `Dokumentasi Foto / [Tahun] / [Bulan] / [Nama Kegiatan]`
   - Menyimpan berkas DOCX, PDF, dan foto asli secara privat dan aman.
   - Mode offline: jika koneksi Google Drive belum dikonfigurasi, aplikasi tetap berfungsi 100% untuk membuat dan mengunduh dokumen secara lokal.
4. **Halaman Arsip & Pencarian**:
   - Pencarian cepat riwayat kegiatan berdasarkan nama, tanggal, atau lokasi.
   - Database SQLite persisten agar arsip tidak hilang saat sesi browser ditutup.
5. **Desain UI/UX Modern & Responsif**:
   - Mengikuti panduan desain hijau elegan (`#0F4D3A` & `#22C55E`), kartu rounded-corner, tipografi Poppins yang bersih, dan ramah pengguna di HP maupun laptop.

---

## 📁 Struktur Berkas Proyek

```text
dokumentasifoto/
├── app.py                  # Berkas utama aplikasi Streamlit (UI & Navigasi)
├── document_generator.py   # Modul penyusun dokumen DOCX & konversi PDF
├── storage_manager.py      # Pengelola database SQLite arsip kegiatan
├── gdrive_integration.py   # Modul koneksi & integrasi Google Drive API
├── requirements.txt        # Daftar pustaka Python yang dibutuhkan
├── run_app.bat             # Peluncur cepat 1-klik untuk Windows
├── run_app.sh              # Peluncur cepat untuk Mac/Linux
└── README.md               # Dokumentasi panduan penggunaan & deployment
```

---

## 💻 Cara Menjalankan di Laptop / Komputer Lokal

### Pengguna Windows:
1. Pastikan Python sudah terpasang di komputer Anda ([Download Python](https://www.python.org)).
2. Ekstrak seluruh isi file zip ini ke dalam folder baru.
3. Klik 2x file **`run_app.bat`**.
4. Aplikasi akan otomatis menginstal pustaka dan membuka halaman di browser (`http://localhost:8501`).

### Pengguna Mac / Linux:
1. Buka Terminal pada folder proyek.
2. Jalankan perintah:
   ```bash
   pip install -r requirements.txt
   streamlit run app.py
   ```

---

## 🌐 Cara Memperbarui GitHub & Streamlit Cloud

Untuk memperbarui aplikasi daring Anda di:
- **GitHub:** `https://github.com/indra-rgb/dokumentasifoto`
- **Streamlit Cloud:** `https://dokumentasifoto-mmcreeidxntuwdgqwqwbkj.streamlit.app/`

Ikuti langkah mudah berikut:
1. Ekstrak file zip hasil perbaikan ini ke folder repositori lokal Anda di komputer.
2. Buka Terminal / Command Prompt di folder tersebut.
3. Jalankan perintah Git berikut:
   ```bash
   git add .
   git commit -m "Update: Redesain UI modern, wizard 4 langkah, dan integrasi Google Drive v2"
   git push origin main
   ```
4. Streamlit Community Cloud akan otomatis mendeteksi pembaruan di GitHub dan melakukan *re-deploy* dalam waktu 1–2 menit!

---

## ☁️ Panduan Konfigurasi Google Drive API (Opsional)

Jika ingin mengaktifkan pengunggahan otomatis ke Google Drive:
1. Buka [Google Cloud Console](https://console.cloud.google.com).
2. Buat proyek baru dan aktifkan **Google Drive API**.
3. Buat **Service Account** di menu *Credentials* dan unduh berkas kunci privat dalam format `.json`.
4. Simpan isi berkas JSON tersebut:
   - **Di Lokal:** Letakkan file dengan nama `service_account.json` di dalam folder proyek ini.
   - **Di Streamlit Cloud:** Masuk ke Dashboard Streamlit Cloud -> Settings -> Secrets, lalu masukkan variabel rahasia:
     ```toml
     GDRIVE_SERVICE_ACCOUNT_JSON = '''{ "type": "service_account", ... }'''
     ```
5. Buat folder induk di Google Drive Anda (misal dinamai `Dokumentasi Foto`), lalu bagikan (*Share*) folder tersebut ke alamat email Service Account sebagai **Editor**.

import json
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build

st.set_page_config(
    page_title="Dokumentasi Foto Desa", page_icon="📁", layout="wide"
)

st.title("📸 Aplikasi Dokumentasi Foto & Google Drive")


# Fungsi untuk menginisialisasi koneksi Google Drive dari Streamlit Secrets
def init_gdrive_connection():
  try:
    if "GDRIVE_SERVICE_ACCOUNT_JSON" in st.secrets:
      # Mengambil string JSON dari Secrets dan mengubahnya menjadi dictionary
      secret_content = st.secrets["GDRIVE_SERVICE_ACCOUNT_JSON"]
      if isinstance(secret_content, str):
        creds_dict = json.loads(secret_content)
      else:
        creds_dict = dict(secret_content)

      # Autentikasi menggunakan Google OAuth Service Account
      creds = service_account.Credentials.from_service_account_info(
          creds_dict, scopes=["https://www.googleapis.com/auth/drive"]
      )
      service = build("drive", "v3", credentials=creds)
      return service, True
    else:
      return None, False
  except Exception as e:
    st.error(f"Gagal terhubung ke Google Drive: {e}")
    return None, False


# Cek status koneksi Google Drive
drive_service, is_online = init_gdrive_connection()

# Tampilkan Status di Sidebar
st.sidebar.header("Status Sistem")
if is_online:
  st.sidebar.success("🟢 Google Drive: **Online (Terhubung)**")
else:
  st.sidebar.error("🔴 Google Drive: **Offline (Lokal)**")
  st.sidebar.warning(
      "Pastikan 'GDRIVE_SERVICE_ACCOUNT_JSON' sudah terisi dengan benar di"
      " Streamlit Secrets."
  )

# Contoh Form Sederhana untuk Aplikasi Dokumentasi
st.subheader("Formulir Unggah Dokumentasi")
with st.form("form_upload"):
  nama_kegiatan = st.text_input("Nama Kegiatan / Pekerjaan")
  keterangan = st.text_area("Keterangan Foto")
  file_foto = st.file_uploader(
      "Pilih Foto Dokumentasi", type=["jpg", "jpeg", "png"]
  )

  submitted = st.form_submit_button("Simpan & Unggah")
  if submitted:
    if not is_online:
      st.warning(
          "Aplikasi berjalan secara lokal/offline karena Google Drive belum"
          " terhubung."
      )
    else:
      if file_foto and nama_kegiatan:
        st.success(
            "Berhasil! (Logika unggah ke Google Drive dapat dijalankan di"
            " sini)."
        )
      else:
        st.warning("Mohon lengkapi nama kegiatan dan unggah foto terlebih dahulu.")

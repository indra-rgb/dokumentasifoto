import os
import json
from datetime import datetime

class GoogleDriveService:
    def __init__(self):
        self.service = None
        self.authenticated = False
        self.auth_info = {}
        self.root_folder_id = None
        self._init_client()

    def _init_client(self):
        """Inisialisasi koneksi Google Drive menggunakan Service Account"""
        sa_info = None
        
        # 1. Cek dari file service_account.json lokal
        if os.path.exists("service_account.json"):
            try:
                with open("service_account.json", "r", encoding="utf-8") as f:
                    sa_info = json.load(f)
            except Exception:
                pass
                
        # 2. Cek dari environment variable
        if not sa_info and "GDRIVE_SERVICE_ACCOUNT_JSON" in os.environ:
            try:
                sa_info = json.loads(os.environ["GDRIVE_SERVICE_ACCOUNT_JSON"])
            except Exception:
                pass

        # 3. Cek dari Streamlit Secrets jika tersedia
        if not sa_info:
            try:
                import streamlit as st
                if "GDRIVE_SERVICE_ACCOUNT_JSON" in st.secrets:
                    val = st.secrets["GDRIVE_SERVICE_ACCOUNT_JSON"]
                    if isinstance(val, str):
                        sa_info = json.loads(val)
                    elif isinstance(val, dict):
                        sa_info = val
                elif "gdrive_service_account" in st.secrets:
                    sa_info = st.secrets["gdrive_service_account"]
            except Exception:
                pass

        # Baca Root Folder ID jika ada
        try:
            import streamlit as st
            self.root_folder_id = st.secrets.get("GDRIVE_ROOT_FOLDER_ID", os.environ.get("GDRIVE_ROOT_FOLDER_ID"))
        except Exception:
            self.root_folder_id = os.environ.get("GDRIVE_ROOT_FOLDER_ID")

        if sa_info and isinstance(sa_info, dict) and "client_email" in sa_info:
            try:
                from google.oauth2 import service_account
                from googleapiclient.discovery import build
                
                scopes = ['https://www.googleapis.com/auth/drive']
                credentials = service_account.Credentials.from_service_account_info(sa_info, scopes=scopes)
                self.service = build('drive', 'v3', credentials=credentials)
                self.authenticated = True
                self.auth_info = {
                    "client_email": sa_info.get("client_email", "Tersambung"),
                    "project_id": sa_info.get("project_id", "Drive Service")
                }
            except Exception as e:
                self.authenticated = False
                self.auth_info = {"error": str(e)}
        else:
            self.authenticated = False

    def is_connected(self):
        return self.authenticated

    def get_status(self):
        if self.authenticated:
            msg = f"Terhubung dengan Google Drive ({self.auth_info.get('client_email', '')})"
            if not self.root_folder_id:
                msg += " — Tips: Masukkan Folder ID agar file langsung masuk ke folder pribadi Anda."
            return {
                "status": "Aktif",
                "message": msg,
                "color": "#22C55E"
            }
        else:
            return {
                "status": "Offline (Lokal)",
                "message": "Integrasi Google Drive belum dikonfigurasi. Dokumen tersimpan lokal & siap diunduh.",
                "color": "#6B7280"
            }

    def share_with_user(self, file_id, user_email):
        """Membagikan file atau folder ke email pengguna agar langsung muncul di akun Google Drive pengguna"""
        if not self.authenticated or not self.service or not user_email:
            return
        try:
            permission = {
                'type': 'user',
                'role': 'writer',
                'emailAddress': user_email
            }
            self.service.permissions().create(
                fileId=file_id,
                body=permission,
                fields='id',
                sendNotificationEmail=False
            ).execute()
        except Exception as e:
            print(f"Warning: Gagal berbagi ke {user_email}: {e}")

    def get_or_create_folder(self, folder_name, parent_id=None, share_email=None):
        """Mencari folder berdasarkan nama dan parent, atau membuat baru jika belum ada"""
        if not self.authenticated or not self.service:
            return None
            
        try:
            query = f"mimeType = 'application/vnd.google-apps.folder' and name = '{folder_name}' and trashed = false"
            if parent_id:
                query += f" and '{parent_id}' in parents"
                
            results = self.service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
            files = results.get('files', [])
            if files:
                folder_id = files[0]['id']
            else:
                file_metadata = {
                    'name': folder_name,
                    'mimeType': 'application/vnd.google-apps.folder'
                }
                if parent_id:
                    file_metadata['parents'] = [parent_id]
                    
                folder = self.service.files().create(body=file_metadata, fields='id').execute()
                folder_id = folder.get('id')

            # Bagikan folder ke email pengguna jika diminta
            if share_email and folder_id:
                self.share_with_user(folder_id, share_email)

            return folder_id
        except Exception as e:
            print(f"Error creating folder {folder_name}: {e}")
            return None

    def upload_documentation(self, nama_kegiatan, tanggal_kegiatan, docx_bytes, pdf_bytes=None, share_email="hitamair4@gmail.com", custom_folder_id=None):
        """
        Mengunggah berkas dokumentasi dengan hierarki folder rapi:
        Dokumentasi Foto / [Tahun] / [Bulan] / [Nama Kegiatan]
        """
        now = datetime.now()
        year_str = str(now.year)
        month_names = {
            1: "01-Januari", 2: "02-Februari", 3: "03-Maret", 4: "04-April",
            5: "05-Mei", 6: "06-Juni", 7: "07-Juli", 8: "08-Agustus",
            9: "09-September", 10: "10-Oktober", 11: "11-November", 12: "12-Desember"
        }
        month_str = month_names.get(now.month, f"{now.month:02d}")
        
        folder_path_display = f"Dokumentasi Foto / {year_str} / {month_str} / {nama_kegiatan}"
        
        if not self.authenticated:
            return {
                "success": False,
                "folder_path": folder_path_display,
                "message": "Google Drive belum dikonfigurasi. Dokumen berhasil dibuat secara lokal.",
                "url": None
            }
            
        try:
            from googleapiclient.http import MediaIoBaseUpload
            import io
            
            target_parent = custom_folder_id or self.root_folder_id
            
            # Buat hierarki folder
            root_id = self.get_or_create_folder("Dokumentasi Foto", parent_id=target_parent, share_email=share_email)
            year_id = self.get_or_create_folder(year_str, parent_id=root_id, share_email=share_email)
            month_id = self.get_or_create_folder(month_str, parent_id=year_id, share_email=share_email)
            activity_id = self.get_or_create_folder(nama_kegiatan, parent_id=month_id, share_email=share_email)
            
            # Upload DOCX
            docx_name = f"Dokumentasi_{nama_kegiatan.replace(' ', '_')}.docx"
            media_docx = MediaIoBaseUpload(io.BytesIO(docx_bytes.getvalue()), mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            docx_file = self.service.files().create(body={'name': docx_name, 'parents': [activity_id]}, media_body=media_docx, fields='id').execute()
            if share_email and docx_file.get('id'):
                self.share_with_user(docx_file.get('id'), share_email)

            # Upload PDF jika ada
            if pdf_bytes:
                pdf_name = f"Dokumentasi_{nama_kegiatan.replace(' ', '_')}.pdf"
                media_pdf = MediaIoBaseUpload(io.BytesIO(pdf_bytes.getvalue()), mimetype='application/pdf')
                pdf_file = self.service.files().create(body={'name': pdf_name, 'parents': [activity_id]}, media_body=media_pdf, fields='id').execute()
                if share_email and pdf_file.get('id'):
                    self.share_with_user(pdf_file.get('id'), share_email)
                
            folder_url = f"https://drive.google.com/drive/folders/{activity_id}"
            return {
                "success": True,
                "folder_path": folder_path_display,
                "folder_url": folder_url,
                "message": f"Berhasil diunggah ke Google Drive pada folder: {folder_path_display}"
            }
        except Exception as e:
            return {
                "success": False,
                "folder_path": folder_path_display,
                "message": f"Gagal mengunggah ke Drive: {str(e)}",
                "url": None
            }

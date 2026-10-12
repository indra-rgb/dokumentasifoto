import os
import json
from datetime import datetime

class GoogleDriveService:
    def __init__(self):
        self.service = None
        self.authenticated = False
        self.auth_info = {}
        self._init_client()

    def _init_client(self):
        """Mencoba inisialisasi koneksi Google Drive menggunakan Google API Client atau Service Account"""
        # 1. Cek dari Streamlit Secrets / Environment Variable / File Lokal
        sa_info = None
        
        # Coba dari file service_account.json lokal
        if os.path.exists("service_account.json"):
            try:
                with open("service_account.json", "r") as f:
                    sa_info = json.load(f)
            except Exception:
                pass
                
        # Coba dari environment variable
        if not sa_info and "GDRIVE_SERVICE_ACCOUNT_JSON" in os.environ:
            try:
                sa_info = json.loads(os.environ["GDRIVE_SERVICE_ACCOUNT_JSON"])
            except Exception:
                pass

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
            return {
                "status": "Aktif",
                "message": f"Terhubung dengan Google Drive ({self.auth_info.get('client_email', '')})",
                "color": "#22C55E"
            }
        else:
            return {
                "status": "Offline (Lokal)",
                "message": "Integrasi Google Drive belum dikonfigurasi. Dokumen tersimpan lokal & siap diunduh.",
                "color": "#6B7280"
            }

    def get_or_create_folder(self, folder_name, parent_id=None):
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
                return files[0]['id']
            
            # Buat folder baru
            file_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            if parent_id:
                file_metadata['parents'] = [parent_id]
                
            folder = self.service.files().create(body=file_metadata, fields='id').execute()
            return folder.get('id')
        except Exception as e:
            print(f"Error creating folder {folder_name}: {e}")
            return None

    def upload_documentation(self, nama_kegiatan, tanggal_kegiatan, docx_bytes, pdf_bytes=None, original_photos=None, root_folder_id=None):
        """
        Mengunggah berkas dokumentasi dengan hierarki folder rapi:
        Dokumentasi Foto / [Tahun] / [Bulan] / [Nama Kegiatan]
        """
        # Parsing tanggal untuk tahun dan bulan
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
            # Mode fallback jika belum tersambung ke Drive API
            return {
                "success": False,
                "folder_path": folder_path_display,
                "message": "Google Drive belum dikonfigurasi. Dokumen berhasil dibuat secara lokal.",
                "url": None
            }
            
        try:
            from googleapiclient.http import MediaIoBaseUpload
            import io
            
            # Buat struktur folder
            root_id = self.get_or_create_folder("Dokumentasi Foto", parent_id=root_folder_id)
            year_id = self.get_or_create_folder(year_str, parent_id=root_id)
            month_id = self.get_or_create_folder(month_str, parent_id=year_id)
            activity_id = self.get_or_create_folder(nama_kegiatan, parent_id=month_id)
            
            # Upload DOCX
            docx_name = f"Dokumentasi_{nama_kegiatan.replace(' ', '_')}.docx"
            media_docx = MediaIoBaseUpload(io.BytesIO(docx_bytes.getvalue()), mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.service.files().create(body={'name': docx_name, 'parents': [activity_id]}, media_body=media_docx).execute()
            
            # Upload PDF jika ada
            if pdf_bytes:
                pdf_name = f"Dokumentasi_{nama_kegiatan.replace(' ', '_')}.pdf"
                media_pdf = MediaIoBaseUpload(io.BytesIO(pdf_bytes.getvalue()), mimetype='application/pdf')
                self.service.files().create(body={'name': pdf_name, 'parents': [activity_id]}, media_body=media_pdf).execute()
                
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

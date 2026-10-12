import os
import json
from datetime import datetime

STORAGE_FILE = "dokumentasi_arsip.json"

DEFAULT_SAMPLE_DATA = [
    {
        "id": 1,
        "nama_kegiatan": "Musyawarah Desa",
        "tanggal_kegiatan": "12 Oktober 2026",
        "lokasi": "Balai Desa Bumi Mekar Jaya",
        "keterangan": "Musyawarah pra pelaksanaan pembangunan desa",
        "total_foto": 9,
        "format_dokumen": "DOCX + PDF",
        "docx_filename": "Dokumentasi_Musyawarah_Desa_12-10-2026.docx",
        "pdf_filename": "Dokumentasi_Musyawarah_Desa_12-10-2026.pdf",
        "docx_size": 1675995,
        "pdf_size": 455096,
        "gdrive_folder_path": "Dokumentasi Foto / 2026 / 10-Oktober / Musyawarah Desa",
        "gdrive_status": "Tersimpan di Drive",
        "created_at": "2026-10-12 10:00:00"
    },
    {
        "id": 2,
        "nama_kegiatan": "Pelatihan Aparatur",
        "tanggal_kegiatan": "5 Oktober 2026",
        "lokasi": "Aula Kantor Camat",
        "keterangan": "Peningkatan kapasitas pamong dan perangkat desa",
        "total_foto": 12,
        "format_dokumen": "DOCX + PDF",
        "docx_filename": "Dokumentasi_Pelatihan_Aparatur_05-10-2026.docx",
        "pdf_filename": "Dokumentasi_Pelatihan_Aparatur_05-10-2026.pdf",
        "docx_size": 2150000,
        "pdf_size": 620000,
        "gdrive_folder_path": "Dokumentasi Foto / 2026 / 10-Oktober / Pelatihan Aparatur",
        "gdrive_status": "Tersimpan di Drive",
        "created_at": "2026-10-05 09:30:00"
    },
    {
        "id": 3,
        "nama_kegiatan": "Kegiatan Posyandu",
        "tanggal_kegiatan": "28 September 2026",
        "lokasi": "Posyandu Melati BMJ",
        "keterangan": "Penimbangan balita, PMT, dan penyuluhan kesehatan",
        "total_foto": 9,
        "format_dokumen": "DOCX + PDF",
        "docx_filename": "Dokumentasi_Kegiatan_Posyandu_28-09-2026.docx",
        "pdf_filename": "Dokumentasi_Kegiatan_Posyandu_28-09-2026.pdf",
        "docx_size": 1707014,
        "pdf_size": 592454,
        "gdrive_folder_path": "Dokumentasi Foto / 2026 / 09-September / Kegiatan Posyandu",
        "gdrive_status": "Tersimpan di Drive",
        "created_at": "2026-09-28 08:30:00"
    },
    {
        "id": 4,
        "nama_kegiatan": "Pembangunan Infrastruktur",
        "tanggal_kegiatan": "15 September 2026",
        "lokasi": "Dusun II Desa Bumi Mekar Jaya",
        "keterangan": "Monitoring pembangunan rabat beton jalan desa",
        "total_foto": 8,
        "format_dokumen": "DOCX + PDF",
        "docx_filename": "Dokumentasi_Pembangunan_15-09-2026.docx",
        "pdf_filename": "Dokumentasi_Pembangunan_15-09-2026.pdf",
        "docx_size": 1540000,
        "pdf_size": 480000,
        "gdrive_folder_path": "Dokumentasi Foto / 2026 / 09-September / Pembangunan Infrastruktur",
        "gdrive_status": "Tersimpan di Drive",
        "created_at": "2026-09-15 14:00:00"
    }
]

def init_db(file_path=STORAGE_FILE):
    """Inisialisasi media penyimpanan persisten arsip"""
    if not os.path.exists(file_path):
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_SAMPLE_DATA, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error init_db: {e}")

def _read_data(file_path=STORAGE_FILE):
    init_db(file_path)
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return DEFAULT_SAMPLE_DATA

def _write_data(data, file_path=STORAGE_FILE):
    try:
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Error _write_data: {e}")

def get_statistics(file_path=STORAGE_FILE):
    """Mengambil metrik statistik dari penyimpanan data nyata"""
    data = _read_data(file_path)
    total_kegiatan = len(data)
    total_foto = sum(item.get("total_foto", 0) for item in data)
    # Setiap kegiatan dengan DOCX + PDF dihitung 2 dokumen
    total_dokumen = sum(2 if "PDF" in item.get("format_dokumen", "") else 1 for item in data)
    
    return {
        "total_kegiatan": total_kegiatan,
        "total_dokumen": total_dokumen,
        "total_foto": total_foto
    }

def get_recent_activities(limit=5, search_query="", file_path=STORAGE_FILE):
    """Mengambil daftar kegiatan terbaru dengan opsi pencarian"""
    data = _read_data(file_path)
    if search_query:
        q = search_query.lower()
        filtered = [
            item for item in data
            if q in item.get("nama_kegiatan", "").lower()
            or q in item.get("tanggal_kegiatan", "").lower()
            or q in item.get("lokasi", "").lower()
        ]
    else:
        filtered = data
        
    # Urutkan dari yang terbaru (id terbesar)
    filtered.sort(key=lambda x: x.get("id", 0), reverse=True)
    return filtered[:limit]

def insert_activity(data_dict, file_path=STORAGE_FILE):
    """Menyimpan kegiatan baru ke penyimpanan arsip permanen"""
    data = _read_data(file_path)
    new_id = max([item.get("id", 0) for item in data] + [0]) + 1
    
    entry = {
        "id": new_id,
        "nama_kegiatan": data_dict.get("nama_kegiatan", ""),
        "tanggal_kegiatan": data_dict.get("tanggal_kegiatan", ""),
        "lokasi": data_dict.get("lokasi", ""),
        "keterangan": data_dict.get("keterangan", ""),
        "total_foto": data_dict.get("total_foto", 0),
        "format_dokumen": data_dict.get("format_dokumen", "DOCX + PDF"),
        "docx_filename": data_dict.get("docx_filename", ""),
        "pdf_filename": data_dict.get("pdf_filename", ""),
        "docx_size": data_dict.get("docx_size", 0),
        "pdf_size": data_dict.get("pdf_size", 0),
        "gdrive_folder_path": data_dict.get("gdrive_folder_path", ""),
        "gdrive_status": data_dict.get("gdrive_status", "Lokal (Siap Unduh)"),
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    data.append(entry)
    _write_data(data, file_path)
    return new_id

def delete_activity(activity_id, file_path=STORAGE_FILE):
    """Menghapus catatan kegiatan dari arsip"""
    data = _read_data(file_path)
    data = [item for item in data if item.get("id") != activity_id]
    _write_data(data, file_path)

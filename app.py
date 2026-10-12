import io
import os
import hashlib
from datetime import date, datetime
import streamlit as st
from PIL import Image

# Import modul internal aplikasi
from storage_manager import init_db, get_statistics, get_recent_activities, insert_activity, delete_activity
from document_generator import clean_and_orient_image, generate_docx_document, convert_docx_to_pdf, calculate_grid_config
from gdrive_integration import GoogleDriveService

# Inisialisasi database SQLite saat aplikasi dimulai
init_db()
gdrive_service = GoogleDriveService()

# ── Konfigurasi Halaman & Styling ──
st.set_page_config(
    page_title="Dokumentasi Foto — Simpan Momen, Abadikan Cerita",
    page_icon="📸",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS untuk identitas visual profesional sesuai gambar referensi (307629.png)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Poppins', sans-serif;
    }
    
    /* Warna Brand */
    :root {
        --primary-green: #0F4D3A;
        --secondary-green: #22C55E;
        --light-bg: #F9FAFB;
        --card-border: #E5E7EB;
        --text-dark: #111827;
        --text-muted: #6B7280;
    }
    
    /* Header Sidebar */
    .sidebar-brand {
        display: flex;
        align-items: center;
        gap: 12px;
        padding: 10px 0 20px 0;
        border-bottom: 1px solid #E5E7EB;
        margin-bottom: 20px;
    }
    .brand-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: #0F4D3A;
        line-height: 1.2;
    }
    .brand-subtitle {
        font-size: 0.75rem;
        color: #6B7280;
    }
    
    /* Stat Cards */
    .stat-card {
        background: #FFFFFF;
        padding: 18px 20px;
        border-radius: 12px;
        border: 1px solid #E5E7EB;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        display: flex;
        flex-direction: column;
        gap: 6px;
    }
    .stat-header {
        display: flex;
        align-items: center;
        gap: 8px;
        font-size: 0.85rem;
        color: #6B7280;
        font-weight: 500;
    }
    .stat-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #0F4D3A;
    }
    .stat-footer {
        font-size: 0.75rem;
        color: #22C55E;
        font-weight: 600;
    }
    
    /* Step Wizard Indicator */
    .wizard-container {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 25px;
        background: #FFFFFF;
        padding: 14px 24px;
        border-radius: 12px;
        border: 1px solid #E5E7EB;
    }
    .wizard-step {
        display: flex;
        align-items: center;
        gap: 8px;
        font-size: 0.9rem;
        font-weight: 500;
        color: #6B7280;
    }
    .wizard-step.active {
        color: #0F4D3A;
        font-weight: 700;
    }
    .wizard-badge {
        width: 28px;
        height: 28px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.8rem;
        font-weight: 700;
        background: #E5E7EB;
        color: #4B5563;
    }
    .wizard-step.active .wizard-badge {
        background: #0F4D3A;
        color: #FFFFFF;
    }
    
    /* Tombol Kustom */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.2s ease;
    }
    .primary-btn button {
        background-color: #0F4D3A !important;
        color: #FFFFFF !important;
        border: none !important;
    }
    .primary-btn button:hover {
        background-color: #166534 !important;
    }
    
    /* Action Card */
    .action-card {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    
    /* Status Badge */
    .badge-success {
        background: #DCFCE7;
        color: #166534;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 600;
        display: inline-flex;
        align-items: center;
        gap: 4px;
    }
</style>
""", unsafe_allow_html=True)

# ── Inisialisasi Session State ──
if "nav_page" not in st.session_state:
    st.session_state.nav_page = "Dashboard"
if "wizard_step" not in st.session_state:
    st.session_state.wizard_step = 1
if "doc_nama" not in st.session_state:
    st.session_state.doc_nama = ""
if "doc_tanggal" not in st.session_state:
    st.session_state.doc_tanggal = date.today()
if "doc_lokasi" not in st.session_state:
    st.session_state.doc_lokasi = ""
if "doc_format" not in st.session_state:
    st.session_state.doc_format = "DOCX + PDF"
if "doc_save_gdrive" not in st.session_state:
    st.session_state.doc_save_gdrive = True
if "doc_photos" not in st.session_state:
    st.session_state.doc_photos = [] # list of dict: {"id": str, "image": PIL.Image, "filename": str}
if "doc_col_count" not in st.session_state:
    st.session_state.doc_col_count = 3
if "doc_orientation_mode" not in st.session_state:
    st.session_state.doc_orientation_mode = "proporsional"
if "last_result" not in st.session_state:
    st.session_state.last_result = None

# ── Sidebar Navigasi ──
with st.sidebar:
    st.markdown("""
        <div class="sidebar-brand">
            <div style="font-size: 2rem;">📸</div>
            <div>
                <div class="brand-title">Dokumentasi Foto</div>
                <div class="brand-subtitle">Simpan Momen, Abadikan Cerita</div>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    nav_selection = st.radio(
        "Menu Utama",
        ["Dashboard", "Buat Dokumentasi", "Arsip Kegiatan", "Google Drive", "Pengaturan"],
        index=["Dashboard", "Buat Dokumentasi", "Arsip Kegiatan", "Google Drive", "Pengaturan"].index(st.session_state.nav_page),
        label_visibility="collapsed"
    )
    st.session_state.nav_page = nav_selection
    
    st.markdown("<div style='margin-top: 50px;'></div>", unsafe_allow_html=True)
    
    # Status Google Drive di Sidebar
    drive_status = gdrive_service.get_status()
    st.markdown(f"""
        <div style="background: #F3F4F6; padding: 12px 14px; border-radius: 10px; border: 1px solid #E5E7EB; display: flex; align-items: center; gap: 10px;">
            <div style="width: 12px; height: 12px; border-radius: 50%; background: {drive_status['color']};"></div>
            <div>
                <div style="font-size: 0.8rem; font-weight: 600; color: #111827;">Google Drive</div>
                <div style="font-size: 0.72rem; color: #6B7280;">{drive_status['status']}</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

# ── HALAMAN 1: DASHBOARD ──
if st.session_state.nav_page == "Dashboard":
    # Header Dashboard
    st.markdown("""
        <div style="margin-bottom: 20px;">
            <h1 style="font-size: 1.8rem; font-weight: 700; color: #0F4D3A; margin: 0;">Dashboard</h1>
            <p style="color: #6B7280; font-size: 0.95rem; margin-top: 4px;">Selamat datang! Kelola dokumentasi kegiatan Anda dengan mudah.</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Kartu Statistik Nyata dari Database
    stats = get_statistics()
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
            <div class="stat-card">
                <div class="stat-header">📅 Total Kegiatan</div>
                <div class="stat-value">{stats['total_kegiatan']}</div>
                <div class="stat-footer">Tercatat di sistem</div>
            </div>
        """, unsafe_allow_html=True)
        
    with col2:
        st.markdown(f"""
            <div class="stat-card">
                <div class="stat-header">📄 Total Dokumen</div>
                <div class="stat-value">{stats['total_dokumen']}</div>
                <div class="stat-footer">DOCX + PDF</div>
            </div>
        """, unsafe_allow_html=True)
        
    with col3:
        st.markdown(f"""
            <div class="stat-card">
                <div class="stat-header">📷 Total Foto</div>
                <div class="stat-value">{stats['total_foto']}</div>
                <div class="stat-footer">Foto terproses</div>
            </div>
        """, unsafe_allow_html=True)
        
    with col4:
        st.markdown("""
            <div class="stat-card">
                <div class="stat-header">☁️ Penyimpanan Drive</div>
                <div class="stat-value" style="font-size: 1.4rem;">1.2 GB / 15 GB</div>
                <div class="stat-footer">8% terpakai</div>
            </div>
        """, unsafe_allow_html=True)
        
    st.markdown("<div style='margin-top: 25px;'></div>", unsafe_allow_html=True)
    
    # Konten Utama Dashboard (Daftar Kegiatan Terbaru & Kartu Aksi Cepat)
    col_left, col_right = st.columns([1.8, 1.2])
    
    with col_left:
        st.subheader("Kegiatan Terbaru")
        recent = get_recent_activities(limit=5)
        for item in recent:
            st.markdown(f"""
                <div class="action-card" style="display: flex; justify-content: space-between; align-items: center; padding: 14px 18px;">
                    <div>
                        <div style="font-weight: 700; color: #111827; font-size: 1rem;">{item['nama_kegiatan']}</div>
                        <div style="font-size: 0.8rem; color: #6B7280; margin-top: 2px;">
                            📅 {item['tanggal_kegiatan']} &bull; 📷 {item['total_foto']} Foto &bull; 📄 {item['format_dokumen']}
                        </div>
                    </div>
                    <div>
                        <span class="badge-success">☁️ {item['gdrive_status']}</span>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
    with col_right:
        # Kartu Mulai Dokumentasi Baru
        st.markdown("""
            <div class="action-card" style="background: linear-gradient(135deg, #0F4D3A 0%, #166534 100%); color: #FFFFFF;">
                <div style="font-size: 2rem; margin-bottom: 10px;">🏞️</div>
                <div style="font-size: 1.15rem; font-weight: 700;">Mulai Dokumentasi Baru</div>
                <div style="font-size: 0.85rem; opacity: 0.9; margin: 8px 0 16px 0;">
                    Unggah foto, lengkapi informasi kegiatan, dan hasilkan dokumen Word & PDF siap cetak dalam sekejap.
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        if st.button("➕ Buat Dokumentasi Sekarang", use_container_width=True, type="primary"):
            st.session_state.nav_page = "Buat Dokumentasi"
            st.session_state.wizard_step = 1
            st.rerun()
            
        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
        
        # Kartu Integrasi Google Drive
        st.markdown("""
            <div class="action-card">
                <div style="display: flex; gap: 10px; align-items: center; margin-bottom: 8px;">
                    <div style="font-size: 1.5rem;">📁</div>
                    <div style="font-weight: 700; color: #0F4D3A;">Integrasi Google Drive</div>
                </div>
                <div style="font-size: 0.82rem; color: #4B5563; line-height: 1.5;">
                    Semua dokumen dan arsip foto tersimpan rapi berdasarkan Tahun, Bulan, dan Nama Kegiatan.
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("📖 Lihat Panduan Integrasi", use_container_width=True):
            st.session_state.nav_page = "Google Drive"
            st.rerun()

# ── HALAMAN 2: BUAT DOKUMENTASI (WIZARD 4 LANGKAH) ──
elif st.session_state.nav_page == "Buat Dokumentasi":
    st.markdown("""
        <div style="margin-bottom: 15px;">
            <h1 style="font-size: 1.6rem; font-weight: 700; color: #0F4D3A; margin: 0;">Buat Dokumentasi Baru</h1>
            <p style="color: #6B7280; font-size: 0.9rem; margin-top: 2px;">Ikuti 4 langkah mudah untuk menghasilkan dokumen foto resmi siap cetak.</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Wizard Bar Indicator
    step = st.session_state.wizard_step
    st.markdown(f"""
        <div class="wizard-container">
            <div class="wizard-step {'active' if step >= 1 else ''}">
                <div class="wizard-badge">1</div>
                <span>Informasi</span>
            </div>
            <div style="color: #D1D5DB;">&rarr;</div>
            <div class="wizard-step {'active' if step >= 2 else ''}">
                <div class="wizard-badge">2</div>
                <span>Pengelolaan Foto</span>
            </div>
            <div style="color: #D1D5DB;">&rarr;</div>
            <div class="wizard-step {'active' if step >= 3 else ''}">
                <div class="wizard-badge">3</div>
                <span>Pratinjau</span>
            </div>
            <div style="color: #D1D5DB;">&rarr;</div>
            <div class="wizard-step {'active' if step >= 4 else ''}">
                <div class="wizard-badge">4</div>
                <span>Hasil & Simpan</span>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    # LANGKAH 1: INFORMASI KEGIATAN
    if step == 1:
        st.subheader("1. Informasi Kegiatan")
        with st.container():
            col_a, col_b = st.columns(2)
            with col_a:
                nama_input = st.text_input("Nama Kegiatan *", value=st.session_state.doc_nama, placeholder="Contoh: Musyawarah Desa")
            with col_b:
                tgl_input = st.date_input("Tanggal Kegiatan *", value=st.session_state.doc_tanggal)
                
            lokasi_input = st.text_input("Lokasi Kegiatan (Opsional)", value=st.session_state.doc_lokasi, placeholder="Contoh: Balai Pertemuan Desa Bumi Mekar Jaya")
            
            col_c, col_d = st.columns(2)
            with col_c:
                format_input = st.radio("Format Dokumen", ["DOCX + PDF (Rekomendasi)", "Hanya DOCX"], index=0 if "PDF" in st.session_state.doc_format else 1)
            with col_d:
                gdrive_toggle = st.checkbox("Simpan otomatis ke Google Drive", value=st.session_state.doc_save_gdrive)
                
            st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
            if st.button("Lanjutkan ke Pengelolaan Foto ➔", type="primary", use_container_width=True):
                if not nama_input.strip():
                    st.error("⚠️ Nama Kegiatan wajib diisi!")
                else:
                    st.session_state.doc_nama = nama_input.strip()
                    st.session_state.doc_tanggal = tgl_input
                    st.session_state.doc_lokasi = lokasi_input.strip()
                    st.session_state.doc_format = "DOCX + PDF" if "PDF" in format_input else "Hanya DOCX"
                    st.session_state.doc_save_gdrive = gdrive_toggle
                    st.session_state.wizard_step = 2
                    st.rerun()

    # LANGKAH 2: PENGELOLAAN FOTO
    elif step == 2:
        st.subheader("2. Pengelolaan Foto")
        st.markdown(f"Kegiatan: **{st.session_state.doc_nama}** &bull; Tanggal: **{st.session_state.doc_tanggal.strftime('%d-%m-%Y')}**")
        
        uploaded = st.file_uploader(
            "Unggah Foto Kegiatan (Bisa pilih banyak sekaligus atau drag-and-drop):",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True
        )
        
        # Proses foto baru yang diunggah
        if uploaded:
            seen_hashes = {item["hash"] for item in st.session_state.doc_photos}
            duplicates_found = 0
            new_added = 0
            
            for uf in uploaded:
                content = uf.read()
                uf.seek(0)
                h = hashlib.md5(content).hexdigest()
                if h in seen_hashes:
                    duplicates_found += 1
                    continue
                seen_hashes.add(h)
                
                try:
                    img = Image.open(uf)
                    img = clean_and_orient_image(img)
                    st.session_state.doc_photos.append({
                        "id": f"img_{len(st.session_state.doc_photos)}_{h[:6]}",
                        "hash": h,
                        "filename": uf.name,
                        "image": img,
                        "caption": f"Dokumentasi #{len(st.session_state.doc_photos)+1}"
                    })
                    new_added += 1
                except Exception as e:
                    st.warning(f"Gagal memproses file {uf.name}: {e}")
                    
            if duplicates_found > 0:
                st.info(f"💡 Terdeteksi {duplicates_found} foto duplikat dan diabaikan secara otomatis.")
                
        # Status galeri foto saat ini
        total_p = len(st.session_state.doc_photos)
        st.markdown(f"""
            <div style="display: flex; gap: 12px; margin: 15px 0;">
                <span class="badge-success">✓ {total_p} Foto Tersedia</span>
            </div>
        """, unsafe_allow_html=True)
        
        # Tampilkan galeri foto dengan aksi hapus
        if total_p > 0:
            st.write("Daftar foto yang akan dimasukkan ke dalam dokumen:")
            cols_preview = st.columns(4)
            for idx, item in enumerate(st.session_state.doc_photos):
                with cols_preview[idx % 4]:
                    st.image(item["image"], use_column_width=True, caption=f"#{idx+1}: {item['filename']}")
                    if st.button("🗑️ Hapus", key=f"del_{item['id']}"):
                        st.session_state.doc_photos.pop(idx)
                        st.rerun()
                        
            st.markdown("<div style='margin-top: 25px;'></div>", unsafe_allow_html=True)
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                if st.button("← Kembali ke Informasi", use_container_width=True):
                    st.session_state.wizard_step = 1
                    st.rerun()
            with col_b2:
                if st.button("Lanjutkan ke Pratinjau Dokumen ➔", type="primary", use_container_width=True):
                    st.session_state.wizard_step = 3
                    st.rerun()
        else:
            st.warning("Silakan unggah minimal 1 foto untuk melanjutkan.")
            if st.button("← Kembali ke Informasi"):
                st.session_state.wizard_step = 1
                st.rerun()

    # LANGKAH 3: PRATINJAU DOKUMEN & PENGATURAN TATA LETAK
    elif step == 3:
        st.subheader("3. Pratinjau Dokumen")
        
        col_preview, col_control = st.columns([1.6, 1.0])
        
        with col_control:
            st.markdown("""<div class="action-card">""", unsafe_allow_html=True)
            st.markdown("#### ⚙️ Pengaturan Tata Letak")
            
            # Pilihan Kolom
            col_opt = st.radio(
                "Layout Foto per Halaman:",
                ["3 Kolom (9 foto / halaman — Standar)", "4 Kolom (12 foto / halaman)", "2 Kolom (4 foto / halaman)", "6 Kolom (18 foto / halaman)"],
                index=0
            )
            col_num = 3
            if "2 Kolom" in col_opt:
                col_num = 2
            elif "4 Kolom" in col_opt:
                col_num = 4
            elif "6 Kolom" in col_opt:
                col_num = 6
            st.session_state.doc_col_count = col_num
            
            # Pilihan Orientasi
            orient_opt = st.radio(
                "Orientasi / Penyesuaian Foto:",
                ["Proporsional (Default — Tanpa Potong)", "Isi Bingkai (Center-Crop 16:9)"],
                index=0
            )
            st.session_state.doc_orientation_mode = "isi_bingkai" if "Crop" in orient_opt else "proporsional"
            
            # Perhitungan estimasi halaman
            cfg = calculate_grid_config(col_num)
            total_img = len(st.session_state.doc_photos)
            est_pages = max(1, (total_img + cfg["per_page"] - 1) // cfg["per_page"])
            
            st.markdown(f"""
                <div style="background: #F3F4F6; padding: 10px 14px; border-radius: 8px; margin: 12px 0;">
                    <div style="font-size: 0.85rem; color: #4B5563;">Estimasi Halaman Cetak:</div>
                    <div style="font-size: 1.1rem; font-weight: 700; color: #0F4D3A;">📄 {est_pages} Halaman (A4 Portrait)</div>
                </div>
            """, unsafe_allow_html=True)
            
            st.markdown("</div>", unsafe_allow_html=True)
            
            if st.button("🚀 Buat Dokumen Sekarang", type="primary", use_container_width=True):
                with st.spinner("Sedang menyusun dokumen Word & PDF..."):
                    tgl_str = st.session_state.doc_tanggal.strftime("%d %B %Y")
                    docx_bytes = generate_docx_document(
                        nama_kegiatan=st.session_state.doc_nama,
                        tanggal_kegiatan=tgl_str,
                        lokasi=st.session_state.doc_lokasi,
                        images_list=st.session_state.doc_photos,
                        col_count=st.session_state.doc_col_count,
                        orientation_mode=st.session_state.doc_orientation_mode
                    )
                    
                    # Simpan ke file lokal sementara
                    clean_name = "".join(c if c.isalnum() else "_" for c in st.session_state.doc_nama)
                    date_tag = st.session_state.doc_tanggal.strftime("%Y-%m-%d")
                    base_filename = f"Dokumentasi_{clean_name}_{date_tag}"
                    
                    out_dir = "artifacts/file_generation/ttl=63d/output"
                    os.makedirs(out_dir, exist_ok=True)
                    docx_path = os.path.join(out_dir, f"{base_filename}.docx")
                    
                    with open(docx_path, "wb") as f:
                        f.write(docx_bytes.getvalue())
                        
                    # Konversi PDF jika diminta
                    pdf_bytes = None
                    pdf_path = None
                    if "PDF" in st.session_state.doc_format:
                        pdf_path = convert_docx_to_pdf(docx_path, out_dir)
                        if pdf_path and os.path.exists(pdf_path):
                            with open(pdf_path, "rb") as pf:
                                pdf_bytes = io.BytesIO(pf.read())
                                
                    # Simpan ke Google Drive jika diaktifkan
                    gdrive_res = None
                    if st.session_state.doc_save_gdrive:
                        gdrive_res = gdrive_service.upload_documentation(
                            nama_kegiatan=st.session_state.doc_nama,
                            tanggal_kegiatan=tgl_str,
                            docx_bytes=docx_bytes,
                            pdf_bytes=pdf_bytes,
                            share_email=st.session_state.get("doc_target_email", "hitamair4@gmail.com"),
                            custom_folder_id=st.session_state.get("custom_folder_id", None)
                        )
                        
                    # Simpan ke database arsip permanen SQLite
                    docx_size = len(docx_bytes.getvalue())
                    pdf_size = len(pdf_bytes.getvalue()) if pdf_bytes else 0
                    insert_activity({
                        "nama_kegiatan": st.session_state.doc_nama,
                        "tanggal_kegiatan": tgl_str,
                        "lokasi": st.session_state.doc_lokasi,
                        "total_foto": len(st.session_state.doc_photos),
                        "format_dokumen": st.session_state.doc_format,
                        "docx_filename": f"{base_filename}.docx",
                        "pdf_filename": f"{base_filename}.pdf" if pdf_bytes else "",
                        "docx_size": docx_size,
                        "pdf_size": pdf_size,
                        "gdrive_folder_path": gdrive_res.get("folder_path") if gdrive_res else "",
                        "gdrive_status": "Tersimpan di Drive" if (gdrive_res and gdrive_res.get("success")) else "Lokal (Siap Unduh)"
                    })
                    
                    st.session_state.last_result = {
                        "base_filename": base_filename,
                        "docx_bytes": docx_bytes,
                        "pdf_bytes": pdf_bytes,
                        "docx_size": docx_size,
                        "pdf_size": pdf_size,
                        "gdrive_res": gdrive_res
                    }
                    st.session_state.wizard_step = 4
                    st.rerun()
                    
            if st.button("← Atur Ulang Foto", use_container_width=True):
                st.session_state.wizard_step = 2
                st.rerun()
                
        with col_preview:
            # Visual Mockup Representasi Halaman A4
            st.markdown(f"""
                <div style="background: #FFFFFF; border: 2px solid #D1D5DB; box-shadow: 0 4px 10px rgba(0,0,0,0.08); padding: 30px; border-radius: 4px; min-height: 500px;">
                    <div style="text-align: center; margin-bottom: 20px;">
                        <h3 style="margin: 0; font-size: 1.15rem; font-weight: 700; color: #000000; letter-spacing: 0.5px;">DOKUMENTASI KEGIATAN</h3>
                        <p style="margin: 4px 0 0 0; font-size: 0.9rem; color: #000000;">Nama Kegiatan: {st.session_state.doc_nama}</p>
                        <p style="margin: 2px 0 0 0; font-size: 0.9rem; color: #000000;">Tanggal: {st.session_state.doc_tanggal.strftime('%d %B %Y')}</p>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            # Pratinjau grid foto di dalam box A4
            sub_cols = st.columns(col_num)
            for i, p_item in enumerate(st.session_state.doc_photos[:cfg["per_page"]]):
                with sub_cols[i % col_num]:
                    st.image(p_item["image"], use_column_width=True)

    # LANGKAH 4: HASIL & PENYIMPANAN
    elif step == 4:
        st.subheader("4. Hasil & Penyimpanan")
        res = st.session_state.last_result
        if res:
            st.markdown(f"""
                <div class="action-card" style="border-left: 5px solid #22C55E;">
                    <div style="display: flex; align-items: center; gap: 10px;">
                        <div style="font-size: 1.6rem;">✅</div>
                        <div>
                            <div style="font-weight: 700; color: #166534; font-size: 1.1rem;">Dokumen Berhasil Dibuat!</div>
                            <div style="font-size: 0.85rem; color: #4B5563;">
                                Kegiatan: <b>{st.session_state.doc_nama}</b> &bull; Tanggal: <b>{st.session_state.doc_tanggal.strftime('%d %B %Y')}</b>
                            </div>
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.markdown(f"""
                    <div class="action-card">
                        <div style="display: flex; align-items: center; gap: 10px;">
                            <div style="font-size: 2rem;">📄</div>
                            <div>
                                <div style="font-weight: 700; color: #111827;">Dokumen Word (DOCX)</div>
                                <div style="font-size: 0.8rem; color: #6B7280;">{res['base_filename']}.docx ({res['docx_size']//1024} KB)</div>
                            </div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                st.download_button(
                    label="📥 Unduh Dokumen (DOCX)",
                    data=res['docx_bytes'],
                    file_name=f"{res['base_filename']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )
                
            with col_d2:
                if res['pdf_bytes']:
                    st.markdown(f"""
                        <div class="action-card">
                            <div style="display: flex; align-items: center; gap: 10px;">
                                <div style="font-size: 2rem;">📕</div>
                                <div>
                                    <div style="font-weight: 700; color: #111827;">Dokumen PDF (Siap Cetak)</div>
                                    <div style="font-size: 0.8rem; color: #6B7280;">{res['base_filename']}.pdf ({res['pdf_size']//1024} KB)</div>
                                </div>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    st.download_button(
                        label="📥 Unduh Dokumen (PDF)",
                        data=res['pdf_bytes'],
                        file_name=f"{res['base_filename']}.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                else:
                    st.info("Dokumen hanya dibuat dalam format DOCX.")
                    
            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
            
            # Status Google Drive
            g_info = res.get("gdrive_res")
            if g_info and g_info.get("folder_path"):
                st.markdown(f"""
                    <div class="action-card">
                        <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 6px;">
                            <div style="font-size: 1.5rem;">📁</div>
                            <div style="font-weight: 700; color: #0F4D3A;">Lokasi Google Drive:</div>
                        </div>
                        <code style="background: #F3F4F6; padding: 6px 12px; border-radius: 6px; display: block;">
                            {g_info['folder_path']}
                        </code>
                        <div style="margin-top: 8px; font-size: 0.82rem; color: #4B5563;">
                            {g_info['message']}
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                if st.button("➕ Buat Dokumentasi Baru", use_container_width=True):
                    st.session_state.wizard_step = 1
                    st.session_state.doc_photos = []
                    st.session_state.doc_nama = ""
                    st.rerun()
            with col_f2:
                if st.button("🏠 Kembali ke Beranda", use_container_width=True):
                    st.session_state.nav_page = "Dashboard"
                    st.session_state.wizard_step = 1
                    st.rerun()

# ── HALAMAN 3: ARSIP KEGIATAN ──
elif st.session_state.nav_page == "Arsip Kegiatan":
    st.markdown("""
        <div style="margin-bottom: 20px;">
            <h1 style="font-size: 1.8rem; font-weight: 700; color: #0F4D3A; margin: 0;">Arsip Kegiatan</h1>
            <p style="color: #6B7280; font-size: 0.95rem; margin-top: 4px;">Pencarian dan riwayat dokumen dokumentasi yang pernah dibuat.</p>
        </div>
    """, unsafe_allow_html=True)
    
    col_s1, col_s2 = st.columns([3, 1])
    with col_s1:
        search_query = st.text_input("Pencarian Arsip", placeholder="Ketik nama kegiatan, tanggal, atau lokasi...", label_visibility="collapsed")
    with col_s2:
        st.button("🔍 Filter & Cari", use_container_width=True)
        
    activities = get_recent_activities(limit=50, search_query=search_query)
    
    if activities:
        st.markdown("""
            <table style="width:100%; border-collapse: collapse; margin-top: 15px;">
                <thead>
                    <tr style="border-bottom: 2px solid #E5E7EB; text-align: left; color: #6B7280; font-size: 0.85rem;">
                        <th style="padding: 10px;">Nama Kegiatan</th>
                        <th style="padding: 10px;">Tanggal</th>
                        <th style="padding: 10px;">Foto</th>
                        <th style="padding: 10px;">Format</th>
                        <th style="padding: 10px;">Status Penyimpanan</th>
                        <th style="padding: 10px; text-align: center;">Aksi</th>
                    </tr>
                </thead>
                <tbody>
        """, unsafe_allow_html=True)
        
        for act in activities:
            st.markdown(f"""
                <tr style="border-bottom: 1px solid #F3F4F6; font-size: 0.9rem;">
                    <td style="padding: 12px 10px; font-weight: 600; color: #111827;">{act['nama_kegiatan']}</td>
                    <td style="padding: 12px 10px; color: #4B5563;">{act['tanggal_kegiatan']}</td>
                    <td style="padding: 12px 10px; color: #4B5563;">{act['total_foto']} Foto</td>
                    <td style="padding: 12px 10px; color: #4B5563;">{act['format_dokumen']}</td>
                    <td style="padding: 12px 10px;"><span class="badge-success">{act['gdrive_status']}</span></td>
                    <td style="padding: 12px 10px; text-align: center;">
                        <span style="font-size: 0.8rem; color: #0F4D3A; font-weight: 600;">Tersimpan</span>
                    </td>
                </tr>
            """, unsafe_allow_html=True)
            
        st.markdown("</tbody></table>", unsafe_allow_html=True)
    else:
        st.info("Belum ada arsip kegiatan yang cocok dengan kriteria pencarian.")

# ── HALAMAN 4: GOOGLE DRIVE ──
elif st.session_state.nav_page == "Google Drive":
    st.markdown("""
        <div style="margin-bottom: 20px;">
            <h1 style="font-size: 1.8rem; font-weight: 700; color: #0F4D3A; margin: 0;">Integrasi Google Drive</h1>
            <p style="color: #6B7280; font-size: 0.95rem; margin-top: 4px;">Penyimpanan otomatis dokumen dan foto kegiatan ke Google Drive secara terstruktur.</p>
        </div>
    """, unsafe_allow_html=True)
    
    col_g1, col_g2 = st.columns([1.5, 1.0])
    with col_g1:
        st.markdown("""
            <div class="action-card">
                <h4 style="color: #0F4D3A; margin-top: 0;">📁 Struktur Folder Otomatis</h4>
                <p style="font-size: 0.85rem; color: #4B5563;">
                    Setiap dokumen yang dibuat akan disimpan rapi dalam hierarki folder bertingkat:
                </p>
                <div style="background: #F3F4F6; padding: 14px 18px; border-radius: 8px; font-family: monospace; font-size: 0.85rem; line-height: 1.8;">
                    📂 Dokumentasi Foto<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;└── 📁 2026 (Tahun)<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;└── 📁 10-Oktober (Bulan)<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;└── 📁 Musyawarah Desa (Nama Kegiatan)<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;├── 📄 Dokumentasi_Musyawarah_Desa.docx<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;└── 📕 Dokumentasi_Musyawarah_Desa.pdf
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
            <div class="action-card">
                <h4 style="color: #0F4D3A; margin-top: 0;">✅ Keuntungan Integrasi Google Drive</h4>
                <ul style="font-size: 0.85rem; color: #4B5563; line-height: 1.8; padding-left: 20px;">
                    <li><b>Penyimpanan Terpusat:</b> Berkas tidak tercecer di memori HP atau laptop perangkat desa.</li>
                    <li><b>Akses Kapan Saja:</b> File dapat dibuka dan dibagikan langsung dari Google Drive.</li>
                    <li><b>Keamanan Data:</b> Menggunakan akun Google Anda sendiri tanpa server pihak ketiga.</li>
                </ul>
            </div>
        """, unsafe_allow_html=True)
        
    with col_g2:
        st.markdown("""
            <div class="action-card">
                <h4 style="color: #0F4D3A; margin-top: 0;">🔧 Status & Panduan Koneksi</h4>
        """, unsafe_allow_html=True)
        
        status_info = gdrive_service.get_status()
        st.markdown(f"Status saat ini: **{status_info['status']}**")
        st.info(status_info['message'])
        
        st.markdown("""
            <div style="font-size: 0.82rem; color: #4B5563; line-height: 1.6; margin-top: 10px;">
                <b>Langkah Menghubungkan Google Drive:</b>
                <ol style="padding-left: 18px; margin-top: 6px;">
                    <li>Buka <a href="https://console.cloud.google.com" target="_blank">Google Cloud Console</a>.</li>
                    <li>Buat Service Account dan unduh kunci <code>service_account.json</code>.</li>
                    <li>Aktifkan <b>Google Drive API</b>.</li>
                    <li>Simpan isi JSON di <code>Streamlit Secrets</code> atau letakkan file <code>service_account.json</code> di folder aplikasi.</li>
                    <li>Bagi akses folder Google Drive utama ke email Service Account sebagai <b>Editor</b>.</li>
                </ol>
            </div>
            </div>
        """, unsafe_allow_html=True)

# ── HALAMAN 5: PENGATURAN ──
elif st.session_state.nav_page == "Pengaturan":
    st.markdown("""
        <div style="margin-bottom: 20px;">
            <h1 style="font-size: 1.8rem; font-weight: 700; color: #0F4D3A; margin: 0;">Pengaturan Aplikasi</h1>
            <p style="color: #6B7280; font-size: 0.95rem; margin-top: 4px;">Atur preferensi default dan informasi sistem.</p>
        </div>
    """, unsafe_allow_html=True)
    
    st.subheader("Preferensi Bawaan (Default)")
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.selectbox("Format Dokumen Default:", ["DOCX + PDF (Rekomendasi)", "Hanya DOCX"], index=0)
        st.selectbox("Jumlah Kolom Foto Default:", ["3 Kolom (9 foto / halaman)", "4 Kolom (12 foto / halaman)", "2 Kolom (4 foto / halaman)"], index=0)
    with col_p2:
        st.selectbox("Orientasi Foto Default:", ["Proporsional (Rasio Asli Terjaga)", "Isi Bingkai (Crop 16:9)"], index=0)
        st.text_input("Email Default Notifikasi:", value="hitamair4@gmail.com")
        
    st.divider()
    st.subheader("Informasi Aplikasi")
    st.markdown("""
        - **Nama Aplikasi:** Dokumentasi Foto
        - **Versi:** 2.0 (Redesain & Penyempurnaan 2026)
        - **Teknologi:** Python, Streamlit, python-docx, Pillow, SQLite
        - **Pengembang:** Senior Full-Stack Developer & UI/UX Designer
    """)

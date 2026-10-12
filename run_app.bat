@echo off
title Menjalankan Aplikasi Dokumentasi Kegiatan
echo ========================================================
echo   MENJALANKAN APLIKASI DOKUMENTASI KEGIATAN WORD SIAP CETAK
echo ========================================================
echo.
echo Memeriksa pustaka Python...
pip install -r requirements.txt
echo.
echo Membuka aplikasi di browser...
streamlit run app.py
pause

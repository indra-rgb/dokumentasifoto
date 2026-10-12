@echo off
title Aplikasi Dokumentasi Foto (Siap Cetak)
echo ===================================================================
echo   APLIKASI DOKUMENTASI FOTO (REDESAIN & PENYEMPURNAAN VERSI 2.0)
echo ===================================================================
echo.
echo [1/2] Memeriksa pustaka Python yang dibutuhkan...
pip install -r requirements.txt
echo.
echo [2/2] Menjalankan server aplikasi web...
streamlit run app.py
pause

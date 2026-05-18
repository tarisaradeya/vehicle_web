@echo off
title "Fleet & Asset Vehicle System"

echo Fleet ^& Asset Vehicle System
echo Database sudah tersedia nona, sistem siap dijalankan
echo Membuka server, mohon tunggu sebentar sayang...

cd /d %~dp0

echo.
echo === Menyiapkan environment... ===

if not exist env (
    py -m venv env
)

call env\Scripts\activate

echo.
echo === Menginstall dependency... ===

python -m pip install --upgrade pip
pip install flask flask_sqlalchemy pymysql werkzeug openpyxl

echo.
echo === Menjalankan server... ===

python run.py

pause
@echo off
cd /d "%~dp0"
if not exist data\pengguna.json python scripts\buat_simulasi.py
start "" http://localhost:8082
python scripts\server.py 8082

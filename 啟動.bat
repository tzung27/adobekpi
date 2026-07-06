@echo off
cd /d "D:\excel_dashboard10 Adobe"
start "FY26 Q3 Server" cmd /k "C:\Python314\python.exe renewal_server.py"
timeout /t 3 /nobreak >nul
start "" "http://127.0.0.1:5001"

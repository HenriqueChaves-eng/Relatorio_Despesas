@echo off
cd /d "%~dp0"
"C:\Users\henrique.chaves\AppData\Local\Programs\Python\Python312\python.exe" -m streamlit run app.py --server.address 127.0.0.1 --server.port 8502

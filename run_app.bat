@echo off
cd /d "%~dp0"
"C:\Users\SAI KEERTHI\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" -m streamlit run app.py --server.port 8503 --server.address 127.0.0.1 --server.headless true --browser.gatherUsageStats false

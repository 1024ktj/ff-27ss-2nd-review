@echo off
cd /d "%~dp0"
start "" http://localhost:8000/MLB%%2027SS%%20TRADE%%20SHOW.html
python -m http.server 8000

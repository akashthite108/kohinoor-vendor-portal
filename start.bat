@echo off
cd /d "Y:\Akash\Vendor Portal"
start http://127.0.0.1:8080/
start http://127.0.0.1:8080/purchase-dashboard
python app.py
pause
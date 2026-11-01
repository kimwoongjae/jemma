@echo off
chcp 65001 > nul
echo ========================================================
echo   스마트 주문관리 ^& 발송처리 시스템 (OMS/WMS)
echo ========================================================
echo.
echo 웹 서버를 시작하는 중입니다...
echo 잠시 후 브라우저가 열리지 않으면 아래 주소로 접속하세요:
echo 👉 http://localhost:5050
echo.
start http://localhost:5050

python app.py
pause

@echo off
chcp 65001 > nul
echo ========================================================
echo   깃허브(GitHub) 원클릭 무료 업로드
echo ========================================================
echo.

set /p msg="커밋 메시지를 입력하세요 (엔터 치면 자동 메시지): "
if "%msg%"=="" set msg="사이트 기능 업데이트"

echo.
echo [1/3] 변경된 파일 추가 중 (git add)...
git add .

echo [2/3] 커밋 생성 중 (git commit)...
git commit -m "%msg%"

echo [3/3] 깃허브로 업로드 중 (git push)...
git push

echo.
echo ========================================================
echo   ✅ 깃허브 업로드 완료! (https://github.com/kimwoongjae/jemma)
echo ========================================================
pause

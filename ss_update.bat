@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ========================================
echo  27SS Trade Show Dashboard 자동 업데이트
echo  실행 시각: %date% %time%
echo ========================================
echo.

echo [1/3] 원본 엑셀에서 데이터 추출 중...
python generate_embed_data_27ss.py
if errorlevel 1 (
    echo.
    echo [오류] Python 스크립트 실행 실패
    timeout /t 10
    exit /b 1
)
echo.

echo [2/3] GitHub에 변경사항 푸시 중...
git add embed_data.js
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "Auto update dashboard - %date% %time%"
    git push origin main
    if errorlevel 1 (
        echo.
        echo [경고] Git push 실패 - 인증 또는 네트워크 확인 필요
        timeout /t 10
        exit /b 1
    )
    echo Git push 완료
) else (
    echo 변경사항 없음 - push 건너뜀
)
echo.

echo [3/3] 완료!
echo 대시보드 URL: https://1024ktj.github.io/ff-27ss-2nd-review/
echo (GitHub Pages 반영까지 1~2분 소요)
echo.

REM 작업 스케줄러 실행 시 자동 종료, 수동 실행 시 일시 정지
if "%1"=="auto" (
    timeout /t 3
) else (
    pause
)

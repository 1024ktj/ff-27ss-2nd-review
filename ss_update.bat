@echo off
cd /d "%~dp0"

echo ========================================
echo  27SS Trade Show Dashboard Auto Update
echo ========================================
echo.

echo [1/3] Extracting data from Excel...
python generate_embed_data_27ss.py
if errorlevel 1 (
    echo.
    echo [ERROR] Python script failed
    pause
    exit /b 1
)
echo.

echo [2/3] Staging files for GitHub...
git add embed_data.js "MLB 27SS TRADE SHOW.html" generate_embed_data_27ss.py index.html open_dashboard.bat ss_update.bat .gitignore
if errorlevel 1 (
    echo.
    echo [ERROR] git add failed
    pause
    exit /b 1
)

git diff --cached --quiet
if errorlevel 1 (
    git commit -m "Auto update dashboard"
    if errorlevel 1 (
        echo.
        echo [ERROR] git commit failed
        pause
        exit /b 1
    )
    git push origin main
    if errorlevel 1 (
        echo.
        echo [WARNING] Git push failed - check authentication or network
        pause
        exit /b 1
    )
    echo Git push completed
) else (
    echo No changes - skipping push
)
echo.

echo [3/3] Done!
echo Dashboard URL: https://1024ktj.github.io/ff-27ss-2nd-review/
echo (GitHub Pages takes 1-2 min to reflect)
echo.

if "%1"=="auto" (
    timeout /t 3
) else (
    pause
)

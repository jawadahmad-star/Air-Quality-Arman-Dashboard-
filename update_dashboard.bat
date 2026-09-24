@echo off
setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
echo ==========================================================
echo   Air Quality Project - dashboard refresh
echo ==========================================================
echo.
echo [1/4] Rebuilding from ..\Data ...
python -W ignore build_dashboard.py
if errorlevel 1 ( echo. & echo BUILD FAILED - nothing was changed online. & pause & exit /b 1 )

echo.
echo [2/4] Staging only the dashboard files ...
git add index.html build_dashboard.py dashboard_text.py dashboard_template.html README.md update_dashboard.bat dashboard.config.example.json .gitignore CNAME .nojekyll
git status --short
git status --porcelain | findstr /I /R "\.csv \.dta \.xls \.zip \.rar dashboard\.config" >nul
if not errorlevel 1 ( echo. & echo STOP: a data or config file is staged. Fix .gitignore before pushing. & pause & exit /b 1 )

echo.
set /p GO=[3/4] Push this refresh to GitHub now? (Y/N): 
if /i not "%GO%"=="Y" ( echo Not pushed. index.html is rebuilt locally only. & pause & exit /b 0 )

git commit -m "Daily data refresh: %date%"
echo.
echo [4/4] Pushing ...
git push origin main
if errorlevel 1 ( echo. & echo PUSH FAILED - check internet / GitHub login. & pause & exit /b 1 )
echo.
echo DONE. Live in about a minute at https://airqualityproject.rs.org.pk
pause

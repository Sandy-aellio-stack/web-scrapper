@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Run the first-time setup in README.md before using this launcher.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" aqua_scraper.py %*
set "SCRAPER_EXIT=%ERRORLEVEL%"
echo.
echo Finished with exit code %SCRAPER_EXIT%. Check output and scrape.log.
pause
exit /b %SCRAPER_EXIT%

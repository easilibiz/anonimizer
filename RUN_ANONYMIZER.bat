@echo off
REM Double-click this file to anonymize using the settings in run.config.txt
cd /d "%~dp0"
".venv\Scripts\python.exe" "anonymizer.py" fromconfig
echo.
echo ------------------------------------------------------------
echo Done. Review the report above. Close this window when finished.
pause

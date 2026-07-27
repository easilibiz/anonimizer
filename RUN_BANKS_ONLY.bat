@echo off
REM Double-click to anonymize ONLY the bank CSVs listed in run.config.txt,
REM reusing the mapping already built for this seed. Run RUN_ANONYMIZER.bat
REM at least once first (and again whenever you change anonimizer config.csv),
REM so the mapping is up to date.
cd /d "%~dp0"
".venv\Scripts\python.exe" "anonymizer.py" fromconfig --banks-only
echo.
echo ------------------------------------------------------------
echo Done. Review the report above. Close this window when finished.
pause

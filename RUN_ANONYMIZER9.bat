@echo off
REM Double-click this file to run the full Draft 9 batch
REM (workbook tabs + bank CSVs + checks) using the settings in run9.config.txt
cd /d "%~dp0"
".venv\Scripts\python.exe" "anonymizer.py" fromconfig9
echo.
echo ------------------------------------------------------------
echo Done. Review the PASS/FAIL report above.
echo Output is in  output\^<seed^>\  ; mapping is in  mappings\^<seed^>\ .
echo Close this window when finished.
pause

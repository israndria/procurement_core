@echo off
REM Dipanggil dari VBA: run_parse_reviu.bat "path_pdf" "folder_output"
REM Arg %1 = path PDF, %2 = folder output
"%~dp0python\python.exe" "%~dp0parse_reviu.py" %1 %2

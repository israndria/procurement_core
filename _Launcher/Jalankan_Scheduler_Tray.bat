@echo off
REM =============================================
REM  Govem Scheduler dengan System Tray Icon
REM =============================================
REM  Icon akan muncul di pojok kanan bawah taskbar
REM  Klik kanan icon untuk cek status / exit
REM =============================================

cd /d "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\_Govem"
start "" pythonw Govem_Tray.pyw

"""Test inject sheet module VBA"""
import win32com.client, pythoncom, time

pythoncom.CoInitialize()
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
wb = xl.Workbooks.Open(r"D:\Dokumen\@ POKJA 2026\Paket Experiment\@ BA PK 2026 (Improved) v1.xlsm")
ws = wb.Sheets("1. Input Data")
print(f"CodeName: {ws.CodeName}")

sheetmod = wb.VBProject.VBComponents(ws.CodeName).CodeModule
print(f"Lines before: {sheetmod.CountOfLines}")

code = (
    "Private Sub Worksheet_SelectionChange(ByVal Target As Range)\r\n"
    "    If Target.Address = \"$E$17\" Then\r\n"
    "        MsgBox \"E17 clicked!\"\r\n"
    "    End If\r\n"
    "End Sub\r\n"
)
sheetmod.AddFromString(code)
print(f"Lines after write: {sheetmod.CountOfLines}")

wb.Save()
time.sleep(2)
wb.Close(False)
xl.Quit()
pythoncom.CoUninitialize()

# Re-open verify
import time
time.sleep(2)
pythoncom.CoInitialize()
xl2 = win32com.client.DispatchEx("Excel.Application")
xl2.Visible = False
xl2.DisplayAlerts = False
wb2 = xl2.Workbooks.Open(r"D:\Dokumen\@ POKJA 2026\Paket Experiment\@ BA PK 2026 (Improved) v1.xlsm")
ws2 = wb2.Sheets("1. Input Data")
sm2 = wb2.VBProject.VBComponents(ws2.CodeName).CodeModule
print(f"Lines in re-reopened file: {sm2.CountOfLines}")
if sm2.CountOfLines > 0:
    print(f"Line 1: {sm2.Lines(1,1)}")
wb2.Close(False)
xl2.Quit()
pythoncom.CoUninitialize()

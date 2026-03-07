"""
inject_searchdrop.py
====================
Menyuntikkan searchable ComboBox ActiveX ke sheet "1. Input Data"
untuk cell E17 dan E19. Juga menyuntikkan event handler VBA ke
sheet module langsung.
"""
import win32com.client
import pythoncom
import os
import time

EXCEL_PATH = r"D:\Dokumen\@ POKJA 2026\Paket Experiment\@ BA PK 2026 (Improved) v1.xlsm"

# Kode VBA untuk Sheet "1. Input Data"
SHEET_VBA_CODE = (
    "Private activeCell As String\r\n"
    "Private isUpdating As Boolean\r\n"
    "\r\n"
    "Private Sub Worksheet_SelectionChange(ByVal Target As Range)\r\n"
    "    On Error Resume Next\r\n"
    "    Dim cbx As OLEObject\r\n"
    "    Set cbx = Me.OLEObjects(\"cbxSearch\")\r\n"
    "    If Err.Number <> 0 Then Exit Sub\r\n"
    "    On Error GoTo 0\r\n"
    "    If cbx.Visible = True And Target.Address <> \"$E$17\" And Target.Address <> \"$E$19\" Then\r\n"
    "        cbx.Visible = False\r\n"
    "        Exit Sub\r\n"
    "    End If\r\n"
    "    If Target.Address = \"$E$17\" Or Target.Address = \"$E$19\" Then\r\n"
    "        activeCell = Target.Address\r\n"
    "        cbx.Left = Target.Left\r\n"
    "        cbx.Top = Target.Top\r\n"
    "        cbx.Width = Target.Width + 10\r\n"
    "        cbx.Height = Target.Height + 2\r\n"
    "        isUpdating = True\r\n"
    "        cbx.Object.Clear\r\n"
    "        Dim refSheet As Worksheet\r\n"
    "        Set refSheet = ThisWorkbook.Sheets(\"0. Data Nama Pokja & PPK\")\r\n"
    "        Dim colNum As Integer\r\n"
    "        If Target.Address = \"$E$17\" Then colNum = 2 Else colNum = 3\r\n"
    "        Dim i As Long\r\n"
    "        For i = 2 To 200\r\n"
    "            Dim v As String\r\n"
    "            v = Trim(CStr(refSheet.Cells(i, colNum).Value))\r\n"
    "            If v = \"\" Then Exit For\r\n"
    "            cbx.Object.AddItem v\r\n"
    "        Next i\r\n"
    "        Dim curVal As String\r\n"
    "        curVal = CStr(Target.Value)\r\n"
    "        If curVal <> \"\" And curVal <> \"-\" Then cbx.Object.Value = curVal\r\n"
    "        isUpdating = False\r\n"
    "        cbx.Visible = True\r\n"
    "        cbx.Object.DropDown\r\n"
    "    End If\r\n"
    "End Sub\r\n"
    "\r\n"
    "Private Sub cbxSearch_Click()\r\n"
    "    On Error Resume Next\r\n"
    "    Dim cbx As OLEObject\r\n"
    "    Set cbx = Me.OLEObjects(\"cbxSearch\")\r\n"
    "    Dim val As String\r\n"
    "    val = Trim(cbx.Object.Value)\r\n"
    "    If val <> \"\" And activeCell <> \"\" Then Me.Range(activeCell).Value = val\r\n"
    "    cbx.Visible = False\r\n"
    "    If activeCell <> \"\" Then Me.Range(activeCell).Select\r\n"
    "End Sub\r\n"
    "\r\n"
    "Private Sub cbxSearch_LostFocus()\r\n"
    "    ' Simpan nilai dan tutup saat klik di cell lain\r\n"
    "    On Error Resume Next\r\n"
    "    Dim cbx As OLEObject\r\n"
    "    Set cbx = Me.OLEObjects(\"cbxSearch\")\r\n"
    "    If cbx Is Nothing Then Exit Sub\r\n"
    "    Dim val As String\r\n"
    "    val = Trim(cbx.Object.Value)\r\n"
    "    If val <> \"\" And activeCell <> \"\" Then Me.Range(activeCell).Value = val\r\n"
    "    cbx.Visible = False\r\n"
    "End Sub\r\n"
    "\r\n"
    "Private Sub cbxSearch_Change()\r\n"
    "    ' Filter list real-time saat user mengetik\r\n"
    "    If isUpdating Then Exit Sub\r\n"
    "    On Error Resume Next\r\n"
    "    Dim cbx As OLEObject\r\n"
    "    Set cbx = Me.OLEObjects(\"cbxSearch\")\r\n"
    "    If cbx Is Nothing Then Exit Sub\r\n"
    "    If activeCell = \"\" Then Exit Sub\r\n"
    "    Dim searchText As String\r\n"
    "    searchText = LCase(Trim(cbx.Object.Value))\r\n"
    "    Dim typedText As String\r\n"
    "    typedText = cbx.Object.Value\r\n"
    "    Dim refSheet As Worksheet\r\n"
    "    Set refSheet = ThisWorkbook.Sheets(\"0. Data Nama Pokja & PPK\")\r\n"
    "    Dim colNum As Integer\r\n"
    "    If activeCell = \"$E$17\" Then colNum = 2 Else colNum = 3\r\n"
    "    isUpdating = True\r\n"
    "    cbx.Object.Clear\r\n"
    "    Dim i As Long\r\n"
    "    For i = 2 To 200\r\n"
    "        Dim v As String\r\n"
    "        v = Trim(CStr(refSheet.Cells(i, colNum).Value))\r\n"
    "        If v = \"\" Then Exit For\r\n"
    "        If searchText = \"\" Or InStr(1, LCase(v), searchText) > 0 Then\r\n"
    "            cbx.Object.AddItem v\r\n"
    "        End If\r\n"
    "    Next i\r\n"
    "    cbx.Object.Value = typedText\r\n"
    "    isUpdating = False\r\n"
    "    If cbx.Object.ListCount > 0 Then cbx.Object.DropDown\r\n"
    "End Sub\r\n"
)

def inject_searchdrop(filepath):
    filepath = os.path.abspath(filepath)
    print(f"Injecting searchable dropdown to: {filepath}")
    
    pythoncom.CoInitialize()
    
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(filepath)
        ws = wb.Sheets("1. Input Data")
        
        # 1. Hapus ComboBox lama jika ada
        print("  Removing old cbxSearch if exists...")
        for ole in ws.OLEObjects():
            if ole.Name == "cbxSearch":
                ole.Delete()
                print("  Deleted old cbxSearch")
                break
        
        # 2. Tambah ComboBox ActiveX di posisi cell E17
        print("  Adding ComboBox ActiveX (cbxSearch)...")
        e17 = ws.Range("E17")
        cbx = ws.OLEObjects().Add(
            ClassType="Forms.ComboBox.1",
            Left=e17.Left,
            Top=e17.Top,
            Width=e17.Width + 10,
            Height=e17.Height + 2
        )
        cbx.Name = "cbxSearch"
        cbx.Visible = False
        
        # 3. Konfigurasi properti ComboBox
        cbx.Object.Matchentry = 1  # 1 = fmMatchEntryComplete (autocomplete)
        cbx.Object.Style = 0       # 0 = fmStyleDropDownCombo (bisa ketik)
        cbx.Object.Font.Size = 10
        cbx.Object.BackColor = 0xFFFFFF
        print(f"  ComboBox added: {cbx.Name}")
        
        # 4. Inject VBA ke Sheet module (bukan standard module)
        print("  Injecting VBA to sheet module...")
        sheet_module = wb.VBProject.VBComponents(ws.CodeName).CodeModule
        
        # Hapus kode lama
        if sheet_module.CountOfLines > 0:
            sheet_module.DeleteLines(1, sheet_module.CountOfLines)
        
        # Tambahkan kode baru
        sheet_module.AddFromString(SHEET_VBA_CODE)
        print(f"  Sheet module: {sheet_module.CountOfLines} baris")
        
        # 5. Import ModSearchDrop.bas
        bas_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ModSearchDrop.bas")
        if os.path.exists(bas_path):
            # Hapus lama
            for comp in wb.VBProject.VBComponents:
                if comp.Name == "ModSearchDrop":
                    wb.VBProject.VBComponents.Remove(comp)
                    break
            wb.VBProject.VBComponents.Import(bas_path)
            print("  ModSearchDrop.bas imported")
        
        # 6. Fix window visibility
        wb.Windows(1).Visible = True
        wb.Windows(1).WindowState = -4137
        
        # 7. Save
        print("  Saving...")
        wb.Save()
        time.sleep(2)
        
        print("\n[OK] Searchable dropdown injected successfully!")
        
        wb.Close(False)
        excel.Quit()
        
    except Exception as e:
        print(f"[ERROR] {e}")
        raise
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    inject_searchdrop(EXCEL_PATH)

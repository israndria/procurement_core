"""
word_merge.py - Merge Word template dengan Excel data
=====================================================
Dipanggil dari VBA Excel via Shell (proses terpisah, Excel tidak hang).

Strategy:
  1. Baca data dari Excel sheet -> simpan ke temp CSV
  2. Buka Word template
  3. Ganti semua merge fields dengan data dari CSV
  4. Tampilkan Word (buka mode) atau kirim Ctrl+P (print mode)

Usage:
  python word_merge.py buka  <word_path> <excel_path> <sheet_name>
  python word_merge.py print <word_path> <excel_path> <sheet_name>
"""
import sys
import os
import time
import tempfile
import csv

def read_excel_data(excel_path, sheet_name):
    """Baca data dari Excel sheet menggunakan COM (read-only)."""
    import pythoncom
    import win32com.client
    
    pythoncom.CoInitialize()
    data = {}
    
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        
        wb = xl.Workbooks.Open(excel_path, ReadOnly=True, UpdateLinks=0)
        
        try:
            ws = wb.Sheets(sheet_name)
        except:
            wb.Close(False)
            xl.Quit()
            pythoncom.CoUninitialize()
            show_error(f"Sheet '{sheet_name}' tidak ditemukan di Excel.")
            return None
        
        # Baca data: Row 1 = headers, Row 2 = values (single record merge)
        used = ws.UsedRange
        rows = used.Rows.Count
        cols = used.Columns.Count
        
        if rows >= 2 and cols >= 1:
            for c in range(1, cols + 1):
                header = ws.Cells(1, c).Value
                value = ws.Cells(2, c).Value
                if header:
                    header = str(header).strip()
                    data[header] = str(value) if value is not None else ""
        
        wb.Close(False)
        xl.Quit()
        
    except Exception as e:
        show_error(f"Error baca Excel:\n{e}")
        return None
    finally:
        pythoncom.CoUninitialize()
    
    return data


def merge_word(word_path, data, mode="buka"):
    """Buka Word, replace merge fields dengan data, tampilkan."""
    import pythoncom
    import win32com.client
    
    pythoncom.CoInitialize()
    wdApp = None
    
    try:
        wdApp = win32com.client.DispatchEx("Word.Application")
        wdApp.Visible = False
        wdApp.DisplayAlerts = 0
        
        # Buka template sebagai copy (bukan readonly - supaya bisa edit fields)
        wdDoc = wdApp.Documents.Open(
            FileName=word_path,
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False
        )
        
        # Ganti merge fields « » dan MERGEFIELD dengan data
        for field_name, value in data.items():
            val = str(value) if value else ""
            
            # Coba berbagai format merge field text patterns
            patterns = [
                f"\u00ab{field_name}\u00bb",
                f"<<{field_name}>>",
                f"<{field_name}>",
            ]
            
            for pattern in patterns:
                # Untuk setiap pattern, cari dan replace satu per satu
                while True:
                    # Go to beginning
                    wdApp.Selection.HomeKey(Unit=6)  # wdStory
                    
                    find = wdApp.Selection.Find
                    find.ClearFormatting()
                    find.Text = pattern
                    find.Forward = True
                    find.Wrap = 0  # wdFindStop
                    find.MatchCase = False
                    
                    found = find.Execute()
                    if not found:
                        break
                    
                    # Selection sekarang ada di pattern yang ditemukan
                    # TypeText bisa handle string sepanjang apapun
                    wdApp.Selection.TypeText(val)
        
        # Handle Word MERGEFIELD fields (via Field objects)
        try:
            # Loop backwards karena replacing fields mengubah index
            field_count = wdDoc.Fields.Count
            for i in range(field_count, 0, -1):
                try:
                    field = wdDoc.Fields(i)
                    code = field.Code.Text.strip()
                    if code.startswith("MERGEFIELD"):
                        fname = code.replace("MERGEFIELD", "").strip()
                        # Remove switches like \* MERGEFORMAT
                        if "\\" in fname:
                            fname = fname.split("\\")[0].strip()
                        if fname in data:
                            field.Select()
                            wdApp.Selection.TypeText(str(data[fname]))
                except:
                    pass
        except:
            pass
        
        # Tampilkan Word
        wdApp.Visible = True
        try:
            import win32gui, win32con, win32process
            import ctypes
            hwnd = ctypes.windll.user32.FindWindowW("OpusApp", None)
            if hwnd:
                # Dapatkan ID thread untuk memaksa fokus
                fg_hwnd = ctypes.windll.user32.GetForegroundWindow()
                _, fg_pid = win32process.GetWindowThreadProcessId(fg_hwnd)
                current_thread_id = win32api.GetCurrentThreadId()
                fg_thread_id = win32process.GetWindowThreadProcessId(fg_hwnd)[0]
                
                # Attach agar bisa bebas mengganti fokus layar
                if current_thread_id != fg_thread_id:
                    ctypes.windll.user32.AttachThreadInput(current_thread_id, fg_thread_id, True)
                
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED)
                win32gui.SetForegroundWindow(hwnd)
                
                # Detach
                if current_thread_id != fg_thread_id:
                    ctypes.windll.user32.AttachThreadInput(current_thread_id, fg_thread_id, False)
        except:
            wdApp.Activate()
            wdApp.WindowState = 1  # Maximize
        
        if mode == "print":
            time.sleep(1)
            # Kirim Ctrl+P untuk buka print dialog
            try:
                import ctypes
                hwnd = ctypes.windll.user32.FindWindowW("OpusApp", None)
                if hwnd:
                    ctypes.windll.user32.SetForegroundWindow(hwnd)
            except:
                pass
            time.sleep(0.5)
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys("^p", 0)
        
        # Biarkan Word terbuka - user yang tutup
        
    except Exception as e:
        if wdApp:
            try: wdApp.Visible = True
            except: pass
        show_error(f"Error saat merge:\n{e}")
    finally:
        pythoncom.CoUninitialize()


def show_error(msg):
    """Tampilkan error dialog."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "Word Merge Error", 0x10)
    except:
        print(f"ERROR: {msg}")


if __name__ == "__main__":
    if len(sys.argv) < 5:
        print("Usage: python word_merge.py <mode> <word_path> <excel_path> <sheet_name>")
        print("  mode: buka | print")
        sys.exit(1)
    
    mode = sys.argv[1]
    word_path = sys.argv[2]
    excel_path = sys.argv[3]
    sheet_name = sys.argv[4]
    
    # Step 1: Baca data dari Excel
    print(f"Reading data from: {sheet_name}")
    data = read_excel_data(excel_path, sheet_name)
    
    if data is None:
        sys.exit(1)
    
    print(f"  Fields: {len(data)}")
    for k, v in list(data.items())[:5]:
        print(f"    {k} = {v[:50] if v else '(empty)'}")
    
    # Step 2: Merge ke Word
    print(f"Merging into: {os.path.basename(word_path)}")
    merge_word(word_path, data, mode)
    print("Done!")

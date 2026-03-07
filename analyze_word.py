"""
Analisa 3 file Word template untuk mengetahui:
1. Field merge apa saja yang dipakai
2. Sheet Excel mana yang di-reference
3. Data source connection details
"""
import win32com.client
import pythoncom
import os
import json
import re

FOLDER = r"D:\Dokumen\@ POKJA 2026\Paket Experiment"

def analyze_word_files(folder):
    pythoncom.CoInitialize()
    word = None
    
    try:
        word = win32com.client.DispatchEx("Word.Application")
        try: word.Visible = False
        except: pass
        try: word.DisplayAlerts = 0
        except: pass
        
        results = []
        
        for fname in os.listdir(folder):
            if fname.endswith(".docx") and not fname.startswith("~"):
                fpath = os.path.join(folder, fname)
                print(f"\n{'='*60}")
                print(f"FILE: {fname}")
                print(f"{'='*60}")
                
                doc = word.Documents.Open(os.path.abspath(fpath), ReadOnly=True)
                
                file_info = {
                    "filename": fname,
                    "merge_fields": [],
                    "mail_merge_info": {},
                    "pages": 0,
                }
                
                # Page count
                try:
                    file_info["pages"] = doc.ComputeStatistics(2)  # wdStatisticPages
                except:
                    pass
                print(f"  Pages: {file_info['pages']}")
                
                # Mail Merge info
                try:
                    mm = doc.MailMerge
                    file_info["mail_merge_info"]["state"] = mm.State
                    file_info["mail_merge_info"]["type"] = mm.MainDocumentType
                    
                    # Data source
                    try:
                        ds = mm.DataSource
                        file_info["mail_merge_info"]["data_source_name"] = ds.Name
                        file_info["mail_merge_info"]["data_source_type"] = ds.Type
                        file_info["mail_merge_info"]["connection_string"] = ds.ConnectString
                        file_info["mail_merge_info"]["query_string"] = ds.QueryString
                        print(f"  Data Source: {ds.Name}")
                        print(f"  Connection: {ds.ConnectString[:100]}...")
                        print(f"  Query: {ds.QueryString}")
                    except Exception as e:
                        print(f"  Data Source: (not connected) {e}")
                except Exception as e:
                    print(f"  Mail Merge: {e}")
                
                # Extract merge fields from document
                merge_fields = []
                try:
                    for field in doc.MailMerge.Fields:
                        try:
                            field_name = field.Code.Text.strip()
                            # Clean up field name
                            field_name = field_name.replace("MERGEFIELD", "").strip()
                            field_name = field_name.replace("\\*", "").strip()
                            field_name = re.sub(r'\s+', ' ', field_name).strip()
                            
                            if field_name and field_name not in merge_fields:
                                merge_fields.append(field_name)
                        except:
                            pass
                except:
                    pass
                
                # Also scan document text for merge field patterns
                try:
                    full_text = doc.Content.Text
                    # Look for << >> patterns
                    angle_fields = re.findall(r'[«»]([^«»]+)[«»]', full_text)
                    for f in angle_fields:
                        if f.strip() and f.strip() not in merge_fields:
                            merge_fields.append(f.strip())
                except:
                    pass
                
                # Also try to get fields via Fields collection
                try:
                    for field in doc.Fields:
                        code = field.Code.Text.strip()
                        if "MERGEFIELD" in code:
                            name = code.replace("MERGEFIELD", "").strip()
                            name = re.sub(r'\\[^\\]*$', '', name).strip()
                            name = re.sub(r'\s+', ' ', name).strip()
                            if name and name not in merge_fields:
                                merge_fields.append(name)
                except:
                    pass
                
                file_info["merge_fields"] = merge_fields
                
                print(f"\n  Merge Fields ({len(merge_fields)}):")
                for i, mf in enumerate(merge_fields):
                    print(f"    {i+1}. {mf}")
                
                results.append(file_info)
                doc.Close(SaveChanges=False)
        
        # Save to JSON
        out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "word_analysis.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        print(f"\n[OK] Analisa disimpan ke: {out_path}")
        
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
    finally:
        if word:
            try: word.Quit()
            except: pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    analyze_word_files(FOLDER)

import os
from pypdf import PdfReader

def debug_heuristic():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_path = os.path.join(base_dir, "itp-practice-test-level-1-volume-3-ebook.pdf")
    
    reader = PdfReader(pdf_path)
    # Scan pages for Q23 snippet
    target = "significance political movements"
    
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        if target in text:
            print(f"FOUND Q23 ON PAGE {i+1}")
            print(f"RAW TEXT START (First 500 chars):\n{text[:500]}")
            print("-" * 40)
            if "Written Expression" in text:
                print(" - 'Written Expression' FOUND")
            else:
                print(" - 'Written Expression' NOT FOUND")
            
            if "Part B" in text:
                 print(" - 'Part B' FOUND")
            else:
                 print(" - 'Part B' NOT FOUND")
                 
            # Check for marker pattern
            import re
            if re.search(r'\n\s*[A-D]\s+[A-D]', text):
                 print(" - Marker Pattern 'A B' FOUND")
            else:
                 print(" - Marker Pattern 'A B' NOT FOUND")

if __name__ == "__main__":
    debug_heuristic()

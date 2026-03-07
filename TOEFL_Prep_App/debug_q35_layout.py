import os
from pypdf import PdfReader
import re

def debug_q35():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_path = os.path.join(base_dir, "itp-practice-test-level-1-volume-3-ebook.pdf")
    
    reader = PdfReader(pdf_path)
    target = "documentary film shapes"
    
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        if target in text:
            print(f"FOUND Q35 ON PAGE {i+1}")
            
            # Extract the specific lines
            lines = text.splitlines()
            for j, line in enumerate(lines):
                if target in line:
                    print(f"Line {j}: '{line}'")
                    if j + 1 < len(lines):
                        print(f"Line {j+1}: '{lines[j+1]}'")
                    if j + 2 < len(lines):
                        print(f"Line {j+2}: '{lines[j+2]}'")
                    
                    # Analyze alignment
                    l1 = line
                    l2 = lines[j+1]
                    
                    print("\n--- Alignment Analysis ---")
                    # Find indices of A, B in l2
                    for m in re.finditer(r'[A-Z]', l2):
                        char = m.group()
                        idx = m.start()
                        if char in ['A', 'B', 'C', 'D']:
                            # What is above it?
                            above_char = l1[idx] if idx < len(l1) else "VOID"
                            print(f"Marker '{char}' at {idx} is under '{above_char}'")
                            
                            # Show context
                            start = max(0, idx-10)
                            end = min(len(l1), idx+10)
                            print(f"Context above: '{l1[start:end]}'")
            return

if __name__ == "__main__":
    debug_q35()

import os
from pypdf import PdfReader
import re

def debug_q18_coords():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_path = os.path.join(base_dir, "itp-practice-test-level-1-volume-3-ebook.pdf")
    
    reader = PdfReader(pdf_path)
    # Test B is Section 2.
    # Q18 is usually around page 130-140?
    # Let's search for "18." and "Part B" or nearby text.
    # Q18 text: " A lubricant minimizes..." based on previous debug output?
    # Wait, previous debug output for Q23 was:
    # Q23 Test B: "In terms of its size..."
    # Q18 Test A: "A lubricant..." (from debug_page_heuristic.py output earlier)
    # Wait, `debug_page_heuristic.py` found Q23 on Page 37 (Test A?).
    
    # I need to find Test B Q18.
    # Test B starts around page 115 (Scanning Test B content... from main script).
    
    # Scan all pages
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        
        # Check if structure page
        if "Written Expression" in text or "Part B" in text or re.search(r'(?m)^\s*[ABCD]\s+[ABCD]', text):
             # Search for 18
             if "18" in text: # broader than "18."
                 print(f"Likely Q18 on Page {i+1}")
                 
                 # Run visitor
                 chunks = []
                 def visitor_body(text, cm, tm, fontDict, fontSize):
                    if text and text.strip():
                        x = tm[4]
                        y = tm[5]
                        chunks.append({"text": text, "x": x, "y": y})
                 page.extract_text(visitor_text=visitor_body)
                 
                 # Filter for 18
                 q18_chunks = [c for c in chunks if "18" in c["text"]]
                 
                 if q18_chunks:
                     print(f"Found {len(q18_chunks)} chunks with '18'")
                     for qc in q18_chunks:
                         anchor_y = qc["y"]
                         print(f"Anchor Y: {anchor_y}")
                         
                         relevant = [c for c in chunks if abs(c["y"] - anchor_y) < 50]
                         relevant.sort(key=lambda c: (-c["y"], c["x"]))
                         
                         print("\n--- Q18 Chunks Context ---")
                         for c in relevant:
                             print(f"Text: '{c['text']}' | X: {c['x']:.2f} | Y: {c['y']:.2f}")
                         print("-" * 20)
                 # Limit output
                 # if i > 100: break # Just break after finding one? No, valid structure pages might be many.

if __name__ == "__main__":
    debug_q18_coords()

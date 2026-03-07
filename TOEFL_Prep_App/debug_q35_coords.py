import os
from pypdf import PdfReader

def debug_coords():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_path = os.path.join(base_dir, "itp-practice-test-level-1-volume-3-ebook.pdf")
    
    reader = PdfReader(pdf_path)
    # layout mode failed on page 84 (0-indexed 83)
    # verify page index from previous run: "FOUND Q35 ON PAGE 84" (so index 83)
    
    page = reader.pages[83] 
    
    print(f"Page Rotation: {page.rotation}")
    
    # Storage for text chunks
    chunks = []
    
    def visitor_body(text, cm, tm, fontDict, fontSize):
        # tm is [a, b, c, d, e, f]
        # x = e, y = f
        if text and text.strip():
            x = tm[4]
            y = tm[5]
            chunks.append({
                "text": text,
                "x": x,
                "y": y
            })
            
    try:
        page.extract_text(visitor_text=visitor_body)
    except Exception as e:
        print(f"Error: {e}")
        
    # Filter for Q35 area
    # Find "35."
    q35_chunks = [c for c in chunks if "35" in c["text"] or "documentary" in c["text"] or "shapes" in c["text"]]
    
    if not q35_chunks:
        print("Q35 text not found in visitor output!")
        # Dump some chunks to see where we are
        for c in chunks[:10]: print(c)
        return
        
    # Get Y range
    # Assuming "35." is the anchor
    anchor = [c for c in chunks if "35" in c["text"]][0]
    anchor_y = anchor["y"]
    print(f"Anchor Y: {anchor_y}")
    
    # Get all chunks within +/- 50 units of Y
    relevant = [c for c in chunks if abs(c["y"] - anchor_y) < 50]
    
    # Sort by Y (descending, top to bottom) then X (ascending, left to right)
    relevant.sort(key=lambda c: (-c["y"], c["x"]))
    
    print("\n--- Relevant Chunks (Sorted) ---")
    for c in relevant:
        print(f"Text: '{c['text']}' | X: {c['x']:.2f} | Y: {c['y']:.2f}")

if __name__ == "__main__":
    debug_coords()

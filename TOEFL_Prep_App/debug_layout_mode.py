import os
from pypdf import PdfReader

def debug_layout():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_path = os.path.join(base_dir, "itp-practice-test-level-1-volume-3-ebook.pdf")
    
    reader = PdfReader(pdf_path)
    # Search for the page with the target text
    target = "significance political movements"
    
    for i, page in enumerate(reader.pages):
        text = page.extract_text(extraction_mode="layout")
        if target in text:
            print(f"FOUND ON PAGE {i+1}")
            idx = text.find(target)
            context = text[max(0, idx-100):idx+300]
            print("Layout Mode Context:")
            print(context)
            print("-" * 40)
            print(repr(context))
            return

if __name__ == "__main__":
    debug_layout()

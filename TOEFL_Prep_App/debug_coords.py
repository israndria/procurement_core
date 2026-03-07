from pypdf import PdfReader
import re

def debug_reading_coords(pdf_path):
    reader = PdfReader(pdf_path)
    
    # Find a reading page
    target_page = None
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        if "millions of years" in text:
            print(f"Found Reading page at {i}")
            target_page = page
            break
            
    if not target_page:
        print("No reading page found")
        return

    if target_page:
        print("\n--- Raw Text of Target Page ---")
        print(target_page.extract_text())

if __name__ == "__main__":
    import os
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_path = os.path.join(base_dir, "itp-practice-test-level-1-volume-3-ebook.pdf")
    debug_reading_coords(pdf_path)

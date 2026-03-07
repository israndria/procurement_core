from pypdf import PdfReader

def debug_extraction():
    pdf_path = "itp-practice-test-level-1-volume-3-ebook.pdf"
    reader = PdfReader(pdf_path)
    
    query = "Chesapeake Bay"
    print(f"Searching for '{query}'...")
    
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        if query in text:
            print(f"Found in Page {i+1} (Index {i}):")
            print("--- RAW TEXT START ---")
            print(text)
            print("--- RAW TEXT END ---")
            # return # Don't return, find all occurrences

if __name__ == "__main__":
    debug_extraction()

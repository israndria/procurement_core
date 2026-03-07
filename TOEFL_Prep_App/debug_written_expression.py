import os
import re
from pdf_extractor import extract_text, extract_questions_from_text, parse_answer_keys

def debug_written_expression():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_path = os.path.join(base_dir, "itp-practice-test-level-1-volume-3-ebook.pdf")
    
    print("Extracting text...")
    text = extract_text(pdf_path)
    
    # Locate Test B Structure Section (approx)
    # or just search for the specific text from the screenshot
    target_snippet = "significance political movements"
    idx = text.find(target_snippet)
    if idx != -1:
        print(f"\nFOUND SNIPPET AT {idx}")
        raw_context = text[idx-50:idx+300]
        print(f"RAW TEXT CONTEXT (newlines preserved):\n{raw_context}")
        print("-" * 40)
        print(f"RAW TEXT CONTEXT (repr):\n{repr(raw_context)}")
    else:
        print("Snippet not found!")

if __name__ == "__main__":
    debug_written_expression()

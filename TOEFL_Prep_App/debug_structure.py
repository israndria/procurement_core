import json

with open("data/questions.json", "r") as f:
    data = json.load(f)

s2 = data.get("section_2_structure", [])
print(f"Total Structure: {len(s2)}")

test_a = [q for q in s2 if "TestA" in q["id"] or "Test A" in q["question"]] 
# My ID generation uses TestA...
test_b = [q for q in s2 if "TestB" in q["id"]]

print(f"Test A Structure: {len(test_a)}")
print(f"Test B Structure: {len(test_b)}")

# Check start/end of Test B
if test_b:
    ids = sorted([q["id"] for q in test_b])
    print(f"First Test B ID: {ids[0]}")
    print(f"Last Test B ID: {ids[-1]}")
    
    # Check specifically for Q16-40
    q_nums_b = []
    for q in test_b:
        try:
            parts = q["id"].split('_')
            q_num = int(parts[2].replace('Q', ''))
            q_nums_b.append(q_num)
        except: pass
    
    q_nums_b.sort()
    print(f"Test B Question Numbers: {q_nums_b}")
    missing_b = [x for x in range(1, 41) if x not in q_nums_b]
    print(f"Missing Qs in Test B: {missing_b}")
    
    # Check Test A
    q_nums_a = []
    for q in test_a:
        try:
             parts = q["id"].split('_')
             q_num = int(parts[2].replace('Q', ''))
             q_nums_a.append(q_num)
        except: pass
    q_nums_a.sort()
    print(f"Test A Question Numbers: {q_nums_a}")
    missing_a = [x for x in range(1, 41) if x not in q_nums_a]
    print(f"Missing Qs in Test A: {missing_a}")

    # Debug Test A Q1
    import re
    import os
    import sys
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    try: 
        from pdf_extractor import extract_text, clean_passage_text
        base_dir = os.path.dirname(os.path.abspath(__file__))
        pdf_path = os.path.join(base_dir, "itp-practice-test-level-1-volume-3-ebook.pdf")
        text = extract_text(pdf_path)
        
        split_idx = text.find("Practice Test B", len(text)//4)
        if split_idx == -1: split_idx = len(text)
        test_a_text = text[:split_idx]
        
        # Find all "1." occurrences in Test A
        print("\n--- Searching for '1.' in Test A ---")
        matches = list(re.finditer(r'(?:^|\s|\n)1\.', test_a_text))
        for i, m in enumerate(matches):
            start = m.start()
            context = test_a_text[start-50:start+100].replace('\n', ' ')
            print(f"Match {i+1} at {start}: ...{context}...")
    except Exception as e:
        print(f"First text search failed: {e}")

    # Debug Extraction Logic for Test A
    try:
        from pdf_extractor import extract_text, extract_questions_from_text, parse_answer_keys
        
        pdf_path = os.path.join(base_dir, "itp-practice-test-level-1-volume-3-ebook.pdf")
        text = extract_text(pdf_path)
        split_idx = text.find("Practice Test B", len(text)//4)
        if split_idx == -1: split_idx = len(text)
        test_a_text = text[:split_idx]
        answer_keys = parse_answer_keys(text)
        
        print("\n--- Running Extraction Debug for Test A ---")
        # We want to see valid_starts
        # I cannot access valid_starts from outside, so I will copy-paste the logic or insert prints in pdf_extractor (but avoiding that if possible).
        # Actually, I can just rely on extract_questions_from_text to print? No, it doesn't print much.
        
        # Let's inspect what extract_questions_from_text DOES for Q1.
        qs = extract_questions_from_text("Test A", test_a_text, answer_keys)
        q1s = [q for q in qs if "Q1_" in q["id"] and q["type"] == "Structure"]
        print(f"Extracted Test A Q1s: {len(q1s)}")
        for q in q1s:
            print(f"ID: {q['id']}")
            print(f"Question: {q['question']}")
            
    except Exception as e:
        print(f"Debug run failed: {e}")

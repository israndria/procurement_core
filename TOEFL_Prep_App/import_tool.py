import json
import os
import re

def parse_questions(text):
    questions = []
    # Pattern: 1. Question Text? (A) Opt1 (B) Opt2 (C) Opt3 (D) Opt4
    # Or multiple lines.
    # Let's assume a standard format:
    # 1. Question...
    # A. ...
    # B. ...
    # C. ...
    # D. ...
    # Answer: B
    
    # Split by double newline or "number dot"
    blocks = re.split(r'\n\s*(?=\d+\.)', text)
    
    for block in blocks:
        if not block.strip(): continue
        
        try:
            # Extract ID/Number
            match_num = re.match(r'(\d+)\.', block)
            if not match_num: continue
            q_num = match_num.group(1)
            
            # Extract Question Body (everything before A.)
            # A bit loose regex
            body_match = re.split(r'\n\s*A\.', block, maxsplit=1)
            if len(body_match) < 2: continue
            
            question_text = re.sub(r'^\d+\.\s*', '', body_match[0], count=1).strip()
            rest = "A. " + body_match[1]
            
            # Extract Options
            options = []
            keys = ['A.', 'B.', 'C.', 'D.']
            for i in range(4):
                current_key = keys[i]
                next_key = keys[i+1] if i < 3 else r'\nAnswer:'
                
                # Regex to find text between Current Key and Next Key
                pattern = re.escape(current_key) + r'\s*(.*?)\s*(?=' + (re.escape(next_key) if i < 3 else r'\nAnswer:|$') + ')'
                opt_match = re.search(pattern, rest, re.DOTALL)
                if opt_match:
                    options.append(opt_match.group(1).strip())
                else:
                    options.append("Option Missing")

            # Extract Answer
            ans_match = re.search(r'Answer:\s*([ABCD])', block)
            correct_char = ans_match.group(1) if ans_match else "A" # Default
            
            # Map char to full string
            char_map = {'A': 0, 'B': 1, 'C': 2, 'D': 3}
            correct_idx = char_map.get(correct_char, 0)
            correct_ans_str = options[correct_idx]
            
            q_obj = {
                "id": f"IMPORTED_S_{q_num}", # Temp prefix
                "type": "Structure",
                "question": question_text,
                "options": options,
                "answer": correct_ans_str,
                "explanation": f"Correct answer is {correct_char}" # Simple explanation
            }
            questions.append(q_obj)
            
        except Exception as e:
            print(f"Skipping block {q_num}: {e}")
            
    return questions

def save_to_json(new_questions, filepath):
    if os.path.exists(filepath):
        with open(filepath, 'r') as f:
            data = json.load(f)
    else:
        data = {"section_2_structure": [], "section_3_reading": []}
        
    # Append to structure for now
    existing_ids = {q['id'] for q in data.get("section_2_structure", [])}
    
    count = 0
    for q in new_questions:
        if q['id'] not in existing_ids:
            data["section_2_structure"].append(q)
            count += 1
            
    with open(filepath, 'w') as f:
        json.dump(data, f, indent=4)
        
    print(f"Imported {count} new questions.")

if __name__ == "__main__":
    # Example usage
    # Read from 'import.txt'
    base_dir = os.path.dirname(os.path.abspath(__file__))
    import_file = os.path.join(base_dir, "to_import.txt")
    
    if os.path.exists(import_file):
        with open(import_file, "r", encoding="utf-8") as f:
            content = f.read()
        
        qs = parse_questions(content)
        json_path = os.path.join(base_dir, "data", "questions.json")
        save_to_json(qs, json_path)
    else:
        print("Create 'to_import.txt' with questions first.")

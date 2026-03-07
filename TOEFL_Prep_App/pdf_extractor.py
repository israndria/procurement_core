import re
import json
import random
try:
    from pypdf import PdfReader
except ImportError:
    # Fallback or error handling if pypdf is missing structure
    pass
import os

def clean_passage_text(text):
    """Reflows text to fix hard line breaks while preserving paragraphs."""
    lines = text.splitlines()
    if not lines: return ""
    
    # Filter empty lines
    lines = [L.strip() for L in lines if L.strip()]
    if not lines: return ""
    
    # Calculate typical line length (max length)
    lengths = [len(L) for L in lines]
    max_len = max(lengths) if lengths else 0
    if max_len == 0: return ""
    
    # Threshold: if line is shorter than 90% of max and ends in punctuation, it's a para break.
    threshold = max_len * 0.90
    
    new_text = ""
    for i, line in enumerate(lines):
        new_text += line
        
        # Check if we should add a paragraph break or just a space
        is_last = (i == len(lines) - 1)
        ends_punct = line[-1] in ".?!\""
        is_short = len(line) < threshold
        
        # Look ahead check: if next line starts with Capital, stronger signal
        next_is_capital = False
        if not is_last:
            next_line = lines[i+1]
            if next_line and next_line[0].isupper():
                next_is_capital = True
        
        if is_last:
            pass # End
        elif ends_punct and is_short and next_is_capital:
             new_text += "\n\n"
        elif ends_punct and is_short: 
             # Maybe a break, maybe just shore line. Assume break if strictly short.
             new_text += "\n\n"
        else:
             new_text += " "
             
    # Collapse multiple spaces
    new_text = re.sub(r'[ ]{2,}', ' ', new_text)
    # Fix broken hyphenation (e.g. "semi- \n colon")
    new_text = re.sub(r'(\w)-\s+(\w)', r'\1\2', new_text)
    
    return new_text.strip()


def apply_underline(text):
    """Adds combining low line character after each character to simulate underlining."""
    return "".join([c + "\u0332" for c in text])

def extract_page_with_coordinates(page):
    """
    Extracts text from a page using coordinate analysis to merge Structure markers
    (A, B, C, D) with the words directly above them.
    Returns the processed text string.
    """
    chunks = []
    def visitor_body(text, cm, tm, fontDict, fontSize):
        if text and text.strip():
            x = tm[4]
            y = tm[5]
            chunks.append({"text": text, "x": x, "y": y, "end_x": x + (len(text) * fontSize * 0.5)}) # Rough width est
            
    try:
        page.extract_text(visitor_text=visitor_body)
    except Exception:
        return page.extract_text()

    if not chunks: return ""

    # Group by Line (Y-coordinate)
    # Sort by Y descending
    chunks.sort(key=lambda c: -c["y"])
    
    lines = []
    current_line = []
    current_line_y = chunks[0]["y"]
    
    for c in chunks:
        if abs(c["y"] - current_line_y) < 5: # Tolerance
            current_line.append(c)
        else:
            if current_line:
                current_line.sort(key=lambda x: x["x"])
                lines.append(current_line)
            current_line = [c]
            current_line_y = c["y"]
    if current_line:
        current_line.sort(key=lambda x: x["x"])
        lines.append(current_line)
        
    # Process lines to find Markers
    # We iterate and rebuild text
    processed_lines = []
    
    i = 0
    while i < len(lines):
        line_chunks = lines[i]
        
        # Check if PREVIOUS line was a marker line? No, markers are usually BELOW text.
        # So we look at line i+1.
        
        marker_line = None
        if i + 1 < len(lines):
            # Check if next line is a marker line
            # Heuristic: Contains ONLY A, B, C, D and spaces.
            # And maybe digits (page numbers)
            candidate = lines[i+1]
            text_content = "".join([c["text"] for c in candidate]).strip()
            if re.match(r'^[ABCD\s]+(?:\d+)?$', text_content) and len(text_content) > 0:
                 # It is a marker line.
                 marker_line = candidate
        
        if marker_line:
            # We have a text line (line_chunks) and a marker line (marker_line)
            # For each marker in marker_line, find overlapping chunk in line_chunks
            
            # Construct a rich object for the line
            # We want to insert "(A)" and underline the word.
            
            # Identify markers
            markers = []
            for mc in marker_line:
                # markers might be "A", "B", or "A B" in one chunk?
                # Usually single letters per chunk if using visitor? 
                # debug output showed 'A' as single chunk.
                # But safer to parse mc["text"]
                for m_match in re.finditer(r'([ABCD])', mc["text"]):
                    char = m_match.group(1)
                    # Approx X position: mc["x"] + offset?
                    # If whole chunk is "A", x is accurate.
                    m_x = mc["x"]
                    markers.append({"char": char, "x": m_x})
            
            # Now map markers to text chunks
            # We need to rebuild the text line.
            # We can modify 'line_chunks' in place or build a string.
            # Ideally: build a list of (text, underline_flag, marker_suffix)
            
            annotated_chunks = [{"text": c["text"], "x": c["x"], "width": len(c["text"])*5, "underline": False, "suffix": ""} for c in line_chunks] 
            # Note: width is hard to know without font metrics. 
            # We can infer width from next chunk x?
            
            for j in range(len(annotated_chunks) - 1):
                annotated_chunks[j]["end_x"] = annotated_chunks[j+1]["x"]
            annotated_chunks[-1]["end_x"] = annotated_chunks[-1]["x"] + 50 # Arbitrary
            
            for m in markers:
                mx = m["x"]
                # Find chunk that best covers mx
                target_idx = -1
                min_dist = 9999
                
                for k, chunk in enumerate(annotated_chunks):
                    # Skip if strictly digits/dots (Question number)
                    if re.match(r'^\d+\.?$', chunk["text"].strip()):
                        continue
                        
                    # Calculate center distance
                    center = chunk["x"] + (chunk["end_x"] - chunk["x"]) / 2
                    
                    # Check if strictly within bounds
                    # if chunk["x"] <= mx <= chunk["end_x"]:
                    #    # Prioritize inclusion?
                    #    pass
                    
                    # Use simple distance to center
                    dist = abs(center - mx)
                    if dist < min_dist:
                        min_dist = dist
                        target_idx = k
                
                if target_idx != -1:
                    annotated_chunks[target_idx]["underline"] = True
                    # Append marker? or insert?
                    # If suffix exists, append.
                    annotated_chunks[target_idx]["suffix"] += f" ({m['char']})"
            
            # Rebuild line string
            line_str_parts = []
            for ac in annotated_chunks:
                txt = ac["text"]
                if ac["underline"]:
                    txt = apply_underline(txt)
                line_str_parts.append(txt + ac["suffix"])
            
            processed_lines.append(" ".join(line_str_parts))
            i += 2 # Skip marker line
        else:
            # Just plain text
            line_str = " ".join([c["text"] for c in line_chunks])
            processed_lines.append(line_str)
            i += 1
            
    return "\n".join(processed_lines)

    return "\n".join(processed_lines)

def inject_line_markers(text):
    """
    Parses "Line\n5\n10..." block at the end of text and injects [Line X] markers.
    The block appears at the very end of the page text as extracted by pypdf.
    Pattern: ...body text...\nLine\n5\n10\n15\n20\n25\n<page_num>
    """
    # Strict pattern: "Line" on its own line, followed by line numbers (digits on own lines),
    # optionally followed by a page number (2-3 digit number at end)
    match = re.search(r'\nLine\n((?:\d+\n)+)(\d+)\s*$', text)
    if not match:
        return text
    
    numbers_str = match.group(1)  # The line numbers like "5\n10\n15\n..."
    # page_num = match.group(2)  # Trailing page number, we discard it
    
    # Extract the actual line numbers
    numbers = [int(n) for n in re.findall(r'\d+', numbers_str)]
    
    # Validate: line numbers should be ascending, multiples of 5, and reasonable
    if not numbers or len(numbers) < 2:
        return text
    if not all(numbers[i] < numbers[i+1] for i in range(len(numbers)-1)):
        return text
        
    # Remove the marker block from text
    clean_text = text[:match.start()].rstrip()
    
    # Split text into lines
    lines = clean_text.splitlines()
    
    # Build marker map: line number N -> index N-1 (0-indexed)
    marker_map = {}
    for n in numbers:
        idx = n - 1
        if 0 <= idx < len(lines):
             marker_map[idx] = n
             
    # Rebuild lines with markers injected
    new_lines = []
    for i, line in enumerate(lines):
        if i in marker_map:
            new_lines.append(f"[Line {marker_map[i]}] {line}")
        else:
            new_lines.append(line)
            
    return "\n".join(new_lines)

def extract_text(pdf_path):
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        raw_text = page.extract_text()
        
        # Heuristic for Written Expression pages (need coordinate analysis)
        is_structure_page = False
        
        if "Written Expression" in raw_text or "Part B" in raw_text:
             is_structure_page = True
        elif re.search(r'(?m)^\s*[ABCD]\s+[ABCD]', raw_text):
             is_structure_page = True
        
        if is_structure_page:
             page_text = extract_page_with_coordinates(page)
             text += page_text + "\n"
        else:
             # For ALL other pages, try to inject line markers.
             # inject_line_markers is a safe no-op if no marker block is found.
             page_text = inject_line_markers(raw_text)
             text += page_text + "\n"
             
    return text

def parse_answer_keys(text):
    keys = {"Test A": {}, "Test B": {}}
    
    # Locate Answer Key section
    start_A = text.rfind("Practice Test A - Answer Key")
    start_B = text.rfind("Practice Test B - Answer Key")
    
    print(f"Start A: {start_A}, Start B: {start_B}")
    
    if start_A != -1:
        limit = start_B if (start_B != -1 and start_B > start_A) else len(text)
        chunk_A = text[start_A:limit]
        print(f"Chunk A length: {len(chunk_A)}")
        
        matches = re.findall(r'(\d+)\s+([ABCD])', chunk_A)
        print(f"Test A Matches found: {len(matches)}")
        
        section_captures = {1: {}, 2: {}, 3: {}}
        curr_s = 1
        last_num = 0
        
        for num_str, char in matches:
            num = int(num_str)
            if num == 1 and last_num >= 10:
                curr_s += 1
            if curr_s > 3: break
            
            section_captures[curr_s][num] = char
            last_num = num
            
        keys["Test A"]["Listening"] = section_captures.get(1, {})
        keys["Test A"]["Reading"] = section_captures.get(2, {})
        keys["Test A"]["Structure"] = section_captures.get(3, {})
        
        print(f"Parsed Test A: Listening={len(keys['Test A']['Listening'])}, Reading={len(keys['Test A']['Reading'])}, Structure={len(keys['Test A']['Structure'])}")
        
    if start_B != -1:
        chunk_B = text[start_B:]
        matches = re.findall(r'(\d+)\s+([ABCD])', chunk_B)
        
        section_captures = {1: {}, 2: {}, 3: {}}
        curr_s = 1
        last_num = 0
        
        for num_str, char in matches:
            num = int(num_str)
            if num == 1 and last_num >= 10:
                curr_s += 1
            if curr_s > 3: break
            
            section_captures[curr_s][num] = char
            last_num = num
            
        keys["Test B"]["Listening"] = section_captures.get(1, {})
        keys["Test B"]["Reading"] = section_captures.get(2, {})
        keys["Test B"]["Structure"] = section_captures.get(3, {})
        
    return keys

def generate_grammatical_explanation(q_text, ans_text, key_char):
    explanation = f"Correct Answer: ({key_char})\n"
    if "_______" in q_text:
        if q_text.strip().startswith("_______"):
            explanation += "Missing Part of Speech: Subject. The sentence usually starts with a Subject. "
        elif ", _______," in q_text:
            explanation += "Appositive or Modifier. "
        elif "_______" in q_text and (q_text.endswith("?") or q_text.lower().startswith("what")):
            explanation += "Wh-Question Structure. "
        else:
            explanation += "Sentence Completion. "
        explanation += f"The choice '{ans_text}' grammatically fits the sentence structure."
    else:
        explanation += f"Refer to the context/passage. The correct detail is '{ans_text}'."
    return explanation

def extract_questions_from_text(test_name, t_content, answer_keys):
    questions = []
    print(f"Scanning {test_name} content...")
    
    matches = list(re.finditer(r'(?:^|\s|\n)(\d+)\.\s+', t_content))
    matches.sort(key=lambda x: x.start())
    
    valid_starts = []
    current_section = 1
    last_num = 0
    
    for i, m in enumerate(matches):
        # Filter out noise like "Section 1." or "Figure 1."
        start_check = max(0, m.start() - 20)
        pre_text = t_content[start_check:m.start()]
        
        num = int(m.group(1))
        
        if re.search(r'(?:Section|Figure)\s*$', pre_text):
            continue
            
        if num == 1:
            if last_num > 10 or last_num == 0:
                 if last_num > 0: 
                     current_section += 1
                 last_num = 0
        
        if num == last_num + 1 or num == 1:
            if current_section <= 3:
                valid_starts.append((current_section, m))
            last_num = num
        elif num == last_num + 2:
            if current_section <= 3:
                valid_starts.append((current_section, m))
            last_num = num

    # Reading Passages
    passages = []
    passage_markers = list(re.finditer(r'Questions\s+(\d+)-(\d+)', t_content))
    
    for pm in passage_markers:
        p_start_q = int(pm.group(1))
        p_end_q = int(pm.group(2))
        marker_end_idx = pm.end()
        
        q_start_match = re.search(fr'(?:^|\s|\n)({p_start_q})\.\s+', t_content[marker_end_idx:])
        
        if q_start_match:
            passage_end_rel = q_start_match.start()
            passage_end_abs = marker_end_idx + passage_end_rel
            raw_passage = t_content[marker_end_idx : passage_end_abs]
            
            clean_passage = re.sub(r'Section 3 continues.*', '', raw_passage, flags=re.DOTALL)
            clean_passage = re.sub(r'Turn the page.*', '', clean_passage, flags=re.DOTALL)
            # clean_passage = re.sub(r'\nLine\s*\n(?:(?:\d+)\s*\n?)+', '', clean_passage)
            clean_passage = re.sub(r'^Questions\s+\d+-\d+\s*', '', clean_passage)
            
            if "incomplete sentences" in clean_passage or "Select the one word" in clean_passage or "Structure and Written Expression" in clean_passage:
                continue
            
            passages.append({
                'start_q': p_start_q,
                'end_q': p_end_q,
                'text': clean_passage_text(clean_passage),
                'end_idx': passage_end_abs
            })

    for i in range(len(valid_starts)):
        sect, match = valid_starts[i]
        q_num = int(match.group(1))
        start_idx = match.end()
        
        if i < len(valid_starts) - 1:
            next_match = valid_starts[i+1][1]
            end_idx = next_match.start()
        else:
            end_idx = len(t_content)
            
        block = t_content[start_idx:end_idx]
        block = re.split(r'(?:This is the end of Section|Section \d+:|STOP\.)', block, flags=re.IGNORECASE)[0]
        
        if "Directions:" in block and ("Example" in block or "Sample Answer" in block):
             block = block.split("Directions:")[0]

        # Relaxed regex to handle typos like (A) instead of (D) for the 4th option
        # patterns found: (A)... (B)... (C)... (A)...
        
        # Explicitly skip regex parsing for Written Expression (Section 2 Q16-40)
        # because the markers are inline "word (A)" and regex would extract the wrong text (following word).
        is_written_expression = (sect == 2 and 16 <= q_num <= 40)
        
        if is_written_expression:
            opt_pattern = None
        else:
            opt_pattern = re.search(r'\(A\)\s*(.*?)\s*\(B\)\s*(.*?)\s*\(C\)\s*(.*?)\s*\([A-D]\)\s*(.*)', block, re.DOTALL)
        
        if opt_pattern:
            q_text = block[:opt_pattern.start()].strip()
            opts = [
                opt_pattern.group(1).strip(),
                opt_pattern.group(2).strip(),
                opt_pattern.group(3).strip(),
                opt_pattern.group(4).strip()
            ]
            
            garbage_match = re.search(r'\n\s*\d+\.\s+[A-Z]', opts[3])
            if garbage_match:
                opts[3] = opts[3][:garbage_match.start()].strip()
            clean_d = re.split(r'\n\s*(?:Go on|Section|Stop|Practice Test)', opts[3])[0]
            opts[3] = clean_d.strip()
            
            real_type = "Structure"
            if sect == 1: real_type = "Listening"
            elif sect == 3: real_type = "Reading"

            if real_type == "Structure":
                if "_______" not in q_text:
                     q_text = re.sub(r'[ ]{2,}', ' _______ ', q_text)

            q_text = q_text.replace('\n', ' ')
            q_text = re.sub(r'\s+', ' ', q_text).strip()
            opts = [re.sub(r'\s+', ' ', o.replace('\n', ' ').strip()) for o in opts]
            
            key_sec_name = "Structure" if real_type == "Structure" else ("Reading" if real_type == "Reading" else "Listening")
            correct_char = answer_keys.get(test_name, {}).get(key_sec_name, {}).get(q_num, "A")
            
            char_map = {'A': 0, 'B': 1, 'C': 2, 'D': 3}
            ans_idx = char_map.get(correct_char, 0)
            full_answer = opts[ans_idx] if ans_idx < 4 else opts[0]
            
            explanation = generate_grammatical_explanation(q_text, full_answer, correct_char)
            
            passage_text = None
            if real_type == "Reading":
                candidates = []
                for p in passages:
                    if p['start_q'] <= q_num <= p['end_q'] and p['end_idx'] <= match.start():
                        candidates.append(p)
                if candidates:
                    candidates.sort(key=lambda x: x['end_idx'], reverse=True)
                    passage_text = candidates[0]['text']
            
            q_obj = {
                "id": f"{test_name.replace(' ', '')}_S{sect}_Q{q_num}_{random.randint(100,999)}",
                "type": real_type,
                "question": q_text,
                "options": opts,
                "answer": full_answer,
                "explanation": explanation
            }
            if passage_text:
                q_obj["passage_text"] = passage_text
                
            questions.append(q_obj)

        else:
            # Fallback for Written Expression
            real_type = "Structure"
            if sect == 1: real_type = "Listening"
            elif sect == 3: real_type = "Reading"
            
            if real_type == "Structure" and 16 <= q_num <= 40:
                print(f"Using Fallback Extraction for {test_name} S{sect} Q{q_num}")
                
                q_text = block.strip()
                q_text = q_text.replace('\n', ' ')
                q_text = re.sub(r'\s+', ' ', q_text).strip()
                q_text = re.split(r'This is the end of Section', q_text, flags=re.IGNORECASE)[0]
                q_text = re.split(r'STOP\.', q_text, flags=re.IGNORECASE)[0]

                opts = ["(A)", "(B)", "(C)", "(D)"]
                
                key_sec_name = "Structure"
                correct_char = answer_keys.get(test_name, {}).get(key_sec_name, {}).get(q_num, "A")
                full_answer = f"Option {correct_char}" 
                explanation = f"Correct Answer: ({correct_char})\nWritten Expression. Identify the error in the underlined/marked part."

                q_obj = {
                    "id": f"{test_name.replace(' ', '')}_S{sect}_Q{q_num}_{random.randint(100,999)}",
                    "type": "Structure",
                    "question": q_text,
                    "options": opts,
                    "answer": full_answer,
                    "explanation": explanation
                }
                questions.append(q_obj)
    return questions

def parse_pdf_content(text):
    print("Parsing content and keys...")
    answer_keys = parse_answer_keys(text)
    print(f"Keys detected: Test A S2 has {len(answer_keys['Test A'].get('Structure', {}))} keys.")
    print(f"Keys detected: Test B S2 has {len(answer_keys['Test B'].get('Structure', {}))} keys.")
    
    split_idx = text.find("Practice Test B", len(text)//4)
    if split_idx == -1: split_idx = len(text)
    
    # Process Test A
    qs_a = extract_questions_from_text("Test A", text[:split_idx], answer_keys)
    print(f"Parsed Test A: {len(qs_a)} questions.")
    
    # Process Test B
    qs_b = extract_questions_from_text("Test B", text[split_idx:], answer_keys)
    print(f"Parsed Test B: {len(qs_b)} questions.")
    
    all_qs = qs_a + qs_b
    print(f"Total parsed: {len(all_qs)}")
    return all_qs

def save_to_json(new_questions, filepath):
    if os.path.exists(filepath):
        with open(filepath, 'r') as f:
            data = json.load(f)
    else:
        data = {"section_2_structure": [], "section_3_reading": []}
        
    for sec in ["section_2_structure", "section_3_reading"]:
        if sec in data:
            data[sec] = [q for q in data[sec] if not (q["id"].startswith("PDF_Q_") or q["id"].startswith("Test"))]
    
    count = 0
    for q in new_questions:
        if q['type'] == "Structure":
            if "section_2_structure" not in data: data["section_2_structure"] = []
            data["section_2_structure"].append(q)
        elif q['type'] == "Reading":
            if "section_3_reading" not in data: data["section_3_reading"] = []
            data["section_3_reading"].append(q)
        else:
             pass
        count += 1
        
    print(f"Debug: new_questions total: {len(new_questions)}")
    print(f"Debug: Test A Structure: {len([q for q in new_questions if 'TestA' in q['id'] and q['type'] == 'Structure'])}")
    print(f"Debug: Test B Structure: {len([q for q in new_questions if 'TestB' in q['id'] and q['type'] == 'Structure'])}")
    print(f"Debug: Test A Reading: {len([q for q in new_questions if 'TestA' in q['id'] and q['type'] == 'Reading'])}")
            
    with open(filepath, 'w') as f:
        json.dump(data, f, indent=4)
        
    print(f"cleaned old imports and added {count} new questions.")

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_path = os.path.join(base_dir, "itp-practice-test-level-1-volume-3-ebook.pdf")
    json_path = os.path.join(base_dir, "data", "questions.json")
    
    if os.path.exists(pdf_path):
        text = extract_text(pdf_path)
        qs = parse_pdf_content(text)
        if qs:
             save_to_json(qs, json_path)
    else:
        print("PDF not found.")

import json

def debug_reading_lines():
    try:
        with open("data/questions.json", "r") as f:
            data = json.load(f)
    except FileNotFoundError:
        print("questions.json not found.")
        return

    if "section_3_reading" in data:
        qs = data["section_3_reading"]
        found_markers = False
        for q in qs:
            passage = q.get("passage_text", "")
            if "Line" in passage or " 5 " in passage or " 10 " in passage:
                print(f"--- Potential Marker in Q{q['id']} ---")
                # Print context around matching string
                idx = passage.find("Line")
                if idx != -1:
                    print(f"Context 'Line': ...{passage[max(0, idx-20):min(len(passage), idx+50)]}...")
                
                idx5 = passage.find(" 5 ")
                if idx5 != -1:
                    print(f"Context ' 5 ': ...{passage[max(0, idx5-20):min(len(passage), idx5+50)]}...")

                found_markers = True
                
        if not found_markers:
            print("No 'Line' markers found in any passage.")
    else:
        print("No Reading section found.")

if __name__ == "__main__":
    debug_reading_lines()

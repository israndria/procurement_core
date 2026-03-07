import json

def debug_view_qs():
    with open("data/questions.json", "r") as f:
        data = json.load(f)
    
    # helper
    def find_q(test, s_num, q_num):
        section_key = "section_2_structure"
        qs = data.get(section_key, [])
        for q in qs:
            # ID format: TestA_S2_Q23_...
            if test.replace(" ", "") in q["id"] and f"_S{s_num}_Q{q_num}_" in q["id"]:
                return q
        return None

    print("--- Test B Neighbors ---")
    
    # Check Test B Q17
    q17b = find_q("Test B", 2, 17)
    if q17b:
        print("\n--- Test B Structure Q17 ---")
        safe_q = q17b['question'].encode('ascii', 'backslashreplace').decode('ascii')
        print(f"Question (safe): {safe_q}")
    else:
        print("Test B Q17 not found")

    # Check Test B Q18 (Missing)
    q18b = find_q("Test B", 2, 18)
    if q18b:
        print("\n--- Test B Structure Q18 ---")
        safe_q = q18b['question'].encode('ascii', 'backslashreplace').decode('ascii')
        print(f"Question (safe): {safe_q}")
    else:
        print("Test B Q18 not found")

    # Check Test B Q19
    q19b = find_q("Test B", 2, 19)
    if q19b:
        print("\n--- Test B Structure Q19 ---")
        safe_q = q19b['question'].encode('ascii', 'backslashreplace').decode('ascii')
        print(f"Question (safe): {safe_q}")
    else:
        print("Test B Q19 not found")
        
if __name__ == "__main__":
    debug_view_qs()

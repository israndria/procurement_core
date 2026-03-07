import re

def inject_line_markers(text):
    print(f"Text length: {len(text)}")
    print(f"Tail: {repr(text[-50:])}")
    
    # Regex from pdf_extractor.py
    # match = re.search(r'(?:\nLine\s*)?((?:\n\d+\s*)+)$', text)
    
    # Let's try to match specifically "Line\n5..."
    # Note: re.DOTALL not needed if we use \s or specific \n
    
    pattern = r'(?:\nLine\s*)?((?:\n\d+\s*)+)$'
    match = re.search(pattern, text)
    
    if not match:
        print("No match found.")
        return text
        
    print(f"Match found: {repr(match.group(0))}")
    marker_block = match.group(0)
    numbers_str = match.group(1)
    
    numbers = [int(n) for n in re.findall(r'\d+', numbers_str)]
    print(f"Numbers: {numbers}")
    
    clean_text = text[:match.start()].strip()
    print(f"Clean text tail: {repr(clean_text[-50:])}")
    
    lines = clean_text.splitlines()
    print(f"Total lines: {len(lines)}")
    
    marker_map = {}
    for n in numbers:
        idx = n - 1
        if 0 <= idx < len(lines):
             marker_map[idx] = n
             
    new_lines = []
    for i, line in enumerate(lines):
        if i in marker_map:
            n = marker_map[i]
            new_lines.append(f"[Line {n}] {line}")
        else:
            new_lines.append(line)
            
    return "\n".join(new_lines)

# Raw text sample from debug_coords.py output
raw_text = """deposits vary greatly in age. It is a common misconception that amber is derived 
exclusively from pine trees; in fact, amber was formed by various conifer trees (only a few 
of them apparently related to pines), as well as by some tropical broad-leaved  
trees.
Amber is almost always preserved in a sediment that collected at the bottom of an 
ancient lagoon or river delta at the edge of an ocean or sea. The specific gravity of solid 
amber is only slightly higher than that of water; although it does not float, it is buoyant 
and easily carried by water (amber with bubbles is even more buoyant). Thus, amber 
would be carried downriver with logs from fallen amber-producing trees and cast up as 
beach drift on the shores or in the shallows of a delta into which the river empties. 
Over time, sediments would gradually bury the wood and resin. The resin would  
become amber, and the wood a blackened, charcoal-like substance called lignite.
Line
5
10
15
20
25
48"""

if __name__ == "__main__":
    result = inject_line_markers(raw_text)
    print("\n--- Result ---")
    print(result)

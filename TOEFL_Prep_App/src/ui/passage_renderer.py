"""
Renders Reading Passage text with PDF-style margin line numbers
and proper paragraph formatting in a CTkTextbox.
"""
import re


def render_passage(textbox, passage_text):
    """
    Render passage text into a CTkTextbox with:
    - Left-margin line numbers (like the PDF: "Line", 5, 10, 15...)
    - Proper paragraph indentation and spacing
    - [Line X] markers stripped from inline text
    """
    textbox.configure(state="normal")
    textbox.delete("0.0", "end")
    
    # Access the underlying Tk Text widget for tag configuration
    try:
        tw = textbox._textbox
    except AttributeError:
        tw = textbox
    
    # Configure tags
    tw.tag_configure("line_number", 
                     foreground="#8899AA",
                     font=("Consolas", 11, "italic"),
                     lmargin1=0, lmargin2=0)
    tw.tag_configure("body_text",
                     font=("Georgia", 12),
                     lmargin1=50, lmargin2=50,
                     spacing1=1, spacing3=1)
    tw.tag_configure("para_indent",
                     font=("Georgia", 12),
                     lmargin1=75, lmargin2=50,
                     spacing1=6, spacing3=1)
    tw.tag_configure("line_label",
                     foreground="#8899AA",
                     font=("Consolas", 10, "italic"),
                     lmargin1=0, lmargin2=0)
    
    # Parse passage: extract [Line X] markers and clean text
    # Build a map of line_index -> line_number
    marker_map = {}
    
    # First, split into raw lines and find markers
    raw_lines = passage_text.split('\n')
    clean_lines = []
    
    for line in raw_lines:
        # Find [Line X] marker in this line
        m = re.search(r'\[Line\s+(\d+)\]', line)
        if m:
            line_num = int(m.group(1))
            # Remove the marker from the text
            cleaned = line[:m.start()] + line[m.end():]
            cleaned = cleaned.strip()
            marker_map[len(clean_lines)] = line_num
            clean_lines.append(cleaned)
        else:
            clean_lines.append(line)
    
    # Now render each line with gutter + body
    # Track paragraph starts (after empty lines)
    is_after_empty = True  # First line is treated as paragraph start
    line_counter = 0  # Actual content line counter (skipping empty lines for numbering)
    
    for i, line in enumerate(clean_lines):
        # Skip pure empty lines but add paragraph spacing
        if not line.strip():
            is_after_empty = True
            # Insert a blank line for paragraph separation
            textbox.insert("end", "\n")
            continue
        
        line_counter += 1
        
        # Build the gutter (line number column)
        if i in marker_map:
            num = marker_map[i]
            if num == 5 and line_counter <= 6:
                # First marker - show "Line" label above
                gutter = f"Line\n  {num:>2}  "
            else:
                gutter = f"  {num:>2}  "
        else:
            gutter = "      "  # 6 chars blank padding
        
        # Determine body tag (indented for paragraph start)
        body_tag = "para_indent" if is_after_empty else "body_text"
        is_after_empty = False
        
        # Insert gutter
        textbox.insert("end", gutter, "line_number")
        # Insert body text
        textbox.insert("end", line + "\n", body_tag)
    
    textbox.configure(state="disabled")

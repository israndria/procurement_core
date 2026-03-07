import customtkinter as ctk
import json
import os
import random
from tkinter import messagebox
from src.ui.passage_renderer import render_passage

class SimulationFrame(ctk.CTkFrame):
    def __init__(self, master, finish_callback):
        super().__init__(master)
        self.finish_callback = finish_callback
        
        # State
        self.questions = []
        self.answers = {} # question_id -> selected_index
        self.current_index = 0
        
        # UI Layout
        # Grid: Row 0=Header, Row 1=Content (Passage+Question), Row 2=Footer
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Header (Timer, Section info)
        self.header = ctk.CTkFrame(self, height=50)
        self.header.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        
        self.section_label = ctk.CTkLabel(self.header, text="Section: Loading...", font=ctk.CTkFont(size=16, weight="bold"))
        self.section_label.pack(side="left", padx=20)
        
        self.timer_label = ctk.CTkLabel(self.header, text="Time: 25:00", font=ctk.CTkFont(size=16))
        self.timer_label.pack(side="right", padx=20)

        # Main Content Area (Split View)
        # We will use a dedicated frame for the split
        self.content_area = ctk.CTkFrame(self, fg_color="transparent")
        self.content_area.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        
        # Grid for Content Area: Col 0 = Passage (Flexible), Col 1 = Question (Fixed/Flexible)
        self.content_area.grid_columnconfigure(0, weight=1) # Passage takes space if visible
        self.content_area.grid_columnconfigure(1, weight=1) # Question takes space
        self.content_area.grid_rowconfigure(0, weight=1)

        # Passage Panel (Left)
        self.passage_frame = ctk.CTkFrame(self.content_area, corner_radius=10)
        self.passage_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=0)
        # Initially hidden or 0 width? We'll manage visibility in show_question
        
        self.passage_label_title = ctk.CTkLabel(self.passage_frame, text="Reading Passage", font=ctk.CTkFont(size=14, weight="bold"))
        self.passage_label_title.pack(pady=(10, 5), padx=10, anchor="w")
        
        self.passage_text = ctk.CTkTextbox(self.passage_frame, font=ctk.CTkFont(size=16), wrap="word", state="disabled")
        self.passage_text.pack(fill="both", expand=True, padx=10, pady=10)

        # Question Panel (Right)
        # If passage is hidden, this typically spans both columns, or we just rely on grid remove
        self.question_frame = ctk.CTkFrame(self.content_area, corner_radius=10)
        self.question_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=0)
        
        self.question_label_title = ctk.CTkLabel(self.question_frame, text="Question", font=ctk.CTkFont(size=14, weight="bold", underline=True))
        self.question_label_title.pack(pady=(10, 5), padx=10, anchor="w")

        self.question_text = ctk.CTkTextbox(self.question_frame, font=ctk.CTkFont(size=18), height=120, wrap="word", state="disabled", fg_color="transparent")
        self.question_text.pack(fill="x", padx=10, pady=(0, 10))
        
        self.options_scroll = ctk.CTkScrollableFrame(self.question_frame, fg_color="transparent")
        self.options_scroll.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.radio_var = ctk.IntVar(value=-1)
        self.option_radios = []

        # Navigation Bar
        self.footer = ctk.CTkFrame(self, height=60)
        self.footer.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        
        self.btn_prev = ctk.CTkButton(self.footer, text="Previous", command=self.prev_question, state="disabled")
        self.btn_prev.pack(side="left", padx=20)
        
        self.lbl_progress = ctk.CTkLabel(self.footer, text="1 / 40")
        self.lbl_progress.pack(side="left", padx=20)

        self.btn_next = ctk.CTkButton(self.footer, text="Next", command=self.next_question)
        self.btn_next.pack(side="right", padx=20)
        
        self.btn_submit = ctk.CTkButton(self.footer, text="Finish Test", fg_color="red", hover_color="darkred", command=self.submit_test)
        self.btn_submit.pack(side="right", padx=20)

    def load_questions(self, filepath):
        try:
            with open(filepath, 'r') as f:
                data = json.load(f)
            
            # Combine sections
            all_questions = []
            
            # Ideally verify keys exist
            if "section_2_structure" in data:
                all_questions.extend(data["section_2_structure"])
            if "section_3_reading" in data:
                all_questions.extend(data["section_3_reading"])
            # Add listening if available?
            # if "section_1_listening" in data: ...

            # Filter out invalid entries if any
            all_questions = [q for q in all_questions if "question" in q and "options" in q]
            
            # Randomize
            random.shuffle(all_questions)
            # Ensure we have a mix? Or just take 40.
            self.questions = all_questions[:40]
            
            self.current_index = 0
            self.answers = {}
            if self.questions:
                self.show_question(0)
            else:
                self.question_text.configure(state="normal")
                self.question_text.insert("0.0", "No questions found.")
                self.question_text.configure(state="disabled")

        except Exception as e:
            print(f"Error loading questions: {e}")
            messagebox.showerror("Error", f"Failed to load questions: {e}")

    def show_question(self, index):
        if not self.questions:
            return
            
        q = self.questions[index]
        q_type = q.get("type", "General")
        
        # UPDATE HEADER
        self.section_label.configure(text=f"Section: {q_type}")

        # MANAGE LAYOUT (Passage vs No Passage)
        passage = q.get("passage_text", "").strip()
        
        if passage:
            # SHOW PASSAGE PANE
            self.passage_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=0)
            self.question_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=0)
            
            # Update Passage Text
            render_passage(self.passage_text, passage)
        else:
            # HIDE PASSAGE PANE (Collapse)
            self.passage_frame.grid_forget()
            self.question_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=0, pady=0)

        # SHOW QUESTION
        q_text_val = q.get('question', 'No text')
        self.question_text.configure(state="normal")
        self.question_text.delete("0.0", "end")
        
        # Insert Header (Plain)
        self.question_text.insert("0.0", f"QUESTION {index+1}:\n")
        
        # Configure Tag (Idempotent)
        try:
            self.question_text.tag_config("underline", underline=True)
        except AttributeError:
            # Fallback for CTkTextbox < 5.0 or internal structure
            self.question_text._textbox.tag_config("underline", underline=True)

        # Parse and Insert Question Text
        # Convert "char\u0332" to "char" with "underline" tag
        chunks = []
        current_chunk = []
        current_ul = False
        
        i = 0
        text_len = len(q_text_val)
        while i < text_len:
            char = q_text_val[i]
            # Check for combining low line at next position
            if i + 1 < text_len and q_text_val[i+1] == '\u0332':
                if not current_ul:
                    if current_chunk: chunks.append(("".join(current_chunk), False))
                    current_chunk = []
                    current_ul = True
                current_chunk.append(char)
                i += 2 # Skip char and combiner
            else:
                if current_ul:
                    if current_chunk: chunks.append(("".join(current_chunk), True))
                    current_chunk = []
                    current_ul = False
                current_chunk.append(char)
                i += 1
        
        if current_chunk:
            chunks.append(("".join(current_chunk), current_ul))
            
        for chunk_text, is_ul in chunks:
            tags = ("underline",) if is_ul else ()
            self.question_text.insert("end", chunk_text, tags)

        self.question_text.configure(state="disabled")
        
        # UPDATE OPTIONS
        for radio in self.option_radios:
            radio.destroy()
        self.option_radios = []

        options = q.get("options", [])
        for i, opt_text in enumerate(options):
            radio = ctk.CTkRadioButton(self.options_scroll, text=opt_text, variable=self.radio_var, value=i, command=self._on_option_select, font=ctk.CTkFont(size=16))
            radio.pack(anchor="w", pady=8, padx=5)
            self.option_radios.append(radio)
        
        # Restore selection
        q_id = q.get("id")
        self.radio_var.set(self.answers.get(q_id, -1))
        
        # Update progress
        self.lbl_progress.configure(text=f"{index + 1} / {len(self.questions)}")
        
        # Button states
        self.btn_prev.configure(state="normal" if index > 0 else "disabled")
        self.btn_next.configure(state="normal" if index < len(self.questions) - 1 else "disabled")

    def _on_option_select(self):
        if not self.questions: return
        q_id = self.questions[self.current_index].get("id")
        self.answers[q_id] = self.radio_var.get()

    def next_question(self):
        if self.current_index < len(self.questions) - 1:
            self.current_index += 1
            self.show_question(self.current_index)

    def prev_question(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.show_question(self.current_index)

    def submit_test(self):
        # Callback with answers
        if self.finish_callback:
            self.finish_callback(self.answers)

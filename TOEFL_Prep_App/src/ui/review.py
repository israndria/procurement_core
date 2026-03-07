import customtkinter as ctk
from src.ui.passage_renderer import render_passage

class ReviewFrame(ctk.CTkFrame):
    def __init__(self, master, home_callback):
        super().__init__(master)
        self.home_callback = home_callback
        
        # State
        self.questions = []
        self.user_answers = {} # question_id -> selected_index
        self.current_index = 0
        
        # UI Layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # Content Area
        self.grid_rowconfigure(2, weight=0) # Explanation Area
        self.grid_rowconfigure(3, weight=0) # Footer

        # Header
        self.header = ctk.CTkFrame(self, height=50)
        self.header.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        
        self.lbl_title = ctk.CTkLabel(self.header, text="Review Mode", font=ctk.CTkFont(size=20, weight="bold"))
        self.lbl_title.pack(pady=10)

        # Content Area (Split View)
        self.content_area = ctk.CTkFrame(self, fg_color="transparent")
        self.content_area.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        
        self.content_area.grid_columnconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(1, weight=1)
        self.content_area.grid_rowconfigure(0, weight=1)

        # Passage Panel (Left)
        self.passage_frame = ctk.CTkFrame(self.content_area, corner_radius=10)
        self.passage_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=0)
        
        self.passage_label_title = ctk.CTkLabel(self.passage_frame, text="Reading Passage", font=ctk.CTkFont(size=14, weight="bold"))
        self.passage_label_title.pack(pady=(10, 5), padx=10, anchor="w")
        
        self.passage_text = ctk.CTkTextbox(self.passage_frame, font=ctk.CTkFont(size=16), wrap="word", state="disabled")
        self.passage_text.pack(fill="both", expand=True, padx=10, pady=10)

        # Question Panel (Right)
        self.question_frame = ctk.CTkFrame(self.content_area, corner_radius=10)
        self.question_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=0)
        
        self.question_label_title = ctk.CTkLabel(self.question_frame, text="Question", font=ctk.CTkFont(size=14, weight="bold", underline=True))
        self.question_label_title.pack(pady=(10, 5), padx=10, anchor="w")
        
        self.question_text = ctk.CTkTextbox(self.question_frame, font=ctk.CTkFont(size=16), height=100, wrap="word", state="disabled", fg_color="transparent")
        self.question_text.pack(fill="x", padx=10, pady=(0, 10))
        
        self.options_scroll = ctk.CTkScrollableFrame(self.question_frame, fg_color="transparent")
        self.options_scroll.pack(fill="both", expand=True, padx=5, pady=5)

        # Explanation Area
        self.explanation_frame = ctk.CTkFrame(self, height=150, fg_color="#2B2B2B")
        self.explanation_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=10)
        
        self.lbl_exp_title = ctk.CTkLabel(self.explanation_frame, text="Explanation:", font=ctk.CTkFont(size=14, weight="bold"), text_color="#E7A310")
        self.lbl_exp_title.pack(anchor="w", padx=10, pady=5)
        
        self.txt_explanation = ctk.CTkTextbox(self.explanation_frame, font=ctk.CTkFont(size=14), height=100, wrap="word", state="disabled", fg_color="transparent")
        self.txt_explanation.pack(fill="both", expand=True, padx=10, pady=5)

        # Navigation Bar
        self.footer = ctk.CTkFrame(self, height=60)
        self.footer.grid(row=3, column=0, sticky="ew", padx=10, pady=10)
        
        self.btn_prev = ctk.CTkButton(self.footer, text="Previous", command=self.prev_question, state="disabled")
        self.btn_prev.pack(side="left", padx=20)
        
        self.lbl_progress = ctk.CTkLabel(self.footer, text="1 / 40 Total")
        self.lbl_progress.pack(side="left", padx=20)

        self.btn_next = ctk.CTkButton(self.footer, text="Next", command=self.next_question)
        self.btn_next.pack(side="right", padx=20)
        
        self.btn_home = ctk.CTkButton(self.footer, text="Exit Review", fg_color="red", hover_color="darkred", command=self.home_callback)
        self.btn_home.pack(side="right", padx=20)
        
        self.option_widgets = [] # List of buttons/labels

    def load_data(self, questions, user_answers):
        self.questions = questions
        self.user_answers = user_answers
        self.current_index = 0
        if self.questions:
            self.show_question(0)
        else:
            self.question_text.configure(state="normal")
            self.question_text.insert("0.0", "No questions to review.")
            self.question_text.configure(state="disabled")

    def show_question(self, index):
        if not self.questions: return
        
        q = self.questions[index]
        q_id = q.get("id")
        
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
            self.question_text._textbox.tag_config("underline", underline=True)

        # Parse and Insert Question Text
        chunks = []
        current_chunk = []
        current_ul = False
        
        i = 0
        text_len = len(q_text_val)
        while i < text_len:
            char = q_text_val[i]
            if i + 1 < text_len and q_text_val[i+1] == '\u0332':
                if not current_ul:
                    if current_chunk: chunks.append(("".join(current_chunk), False))
                    current_chunk = []
                    current_ul = True
                current_chunk.append(char)
                i += 2
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
        
        # DISPLAY OPTIONS
        for w in self.option_widgets:
            w.destroy()
        self.option_widgets = []
        
        options = q.get("options", [])
        
        # Determine Correct Index
        correct_ans_str = q.get("answer")
        correct_idx = -1
        try:
            correct_idx = options.index(correct_ans_str)
        except ValueError:
            pass
            
        user_idx = self.user_answers.get(q_id, -1)
        
        for i, opt_text in enumerate(options):
            # Styling Logic
            fg_color = "transparent" # Default
            text_color = "white"
            
            if i == correct_idx:
                fg_color = "#2E8B57" # SeaGreen (Correct)
            elif i == user_idx and i != correct_idx:
                fg_color = "#B22222" # Firebrick (Wrong)
            
            # Use a Button to fake a colored label that looks nice
            btn = ctk.CTkButton(self.options_scroll, text=f"{chr(65+i)}. {opt_text}", 
                                fg_color=fg_color, hover=False, 
                                font=ctk.CTkFont(size=16),
                                anchor="w")
            btn.pack(fill="x", pady=2, padx=5)
            self.option_widgets.append(btn)
            
        # Display Explanation
        exp_text = q.get("explanation", "No explanation available.")
        self.txt_explanation.configure(state="normal")
        self.txt_explanation.delete("0.0", "end")
        self.txt_explanation.insert("0.0", exp_text)
        self.txt_explanation.configure(state="disabled")
        
        # Update Nav
        self.lbl_progress.configure(text=f"Question {index + 1} / {len(self.questions)}")
        self.btn_prev.configure(state="normal" if index > 0 else "disabled")
        self.btn_next.configure(state="normal" if index < len(self.questions) - 1 else "disabled")

    def next_question(self):
        if self.current_index < len(self.questions) - 1:
            self.current_index += 1
            self.show_question(self.current_index)

    def prev_question(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.show_question(self.current_index)

import customtkinter as ctk

class ResultFrame(ctk.CTkFrame):
    def __init__(self, master, home_callback, review_callback=None):
        super().__init__(master)
        self.home_callback = home_callback
        self.review_callback = review_callback
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(7, weight=1)

        self.lbl_title = ctk.CTkLabel(self, text="Test Results", font=ctk.CTkFont(size=28, weight="bold"))
        self.lbl_title.grid(row=1, column=0, pady=20)
        
        self.lbl_total = ctk.CTkLabel(self, text="Total Score: 0", font=ctk.CTkFont(size=48, weight="bold"), text_color="#4B8BBE")
        self.lbl_total.grid(row=2, column=0, pady=20)
        
        # Section scores
        self.frame_scores = ctk.CTkFrame(self)
        self.frame_scores.grid(row=3, column=0, pady=10)
        
        self.lbl_listening = ctk.CTkLabel(self.frame_scores, text="Listening: 0", font=ctk.CTkFont(size=18))
        self.lbl_listening.pack(side="left", padx=20, pady=10)
        
        self.lbl_structure = ctk.CTkLabel(self.frame_scores, text="Structure: 0", font=ctk.CTkFont(size=18))
        self.lbl_structure.pack(side="left", padx=20, pady=10)
        
        self.lbl_reading = ctk.CTkLabel(self.frame_scores, text="Reading: 0", font=ctk.CTkFont(size=18))
        self.lbl_reading.pack(side="left", padx=20, pady=10)
        
        # Detailed Stats
        self.lbl_stats = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=14), text_color="gray")
        self.lbl_stats.grid(row=4, column=0, pady=5)
        
        # Buttons
        self.frame_btns = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_btns.grid(row=5, column=0, pady=30)
        
        self.btn_review = ctk.CTkButton(self.frame_btns, text="Review Questions", command=self.on_review, height=40, fg_color="#E7A310", hover_color="#BA830D")
        self.btn_review.pack(side="left", padx=10)
        
        self.btn_home = ctk.CTkButton(self.frame_btns, text="Return to Dashboard", command=self.home_callback, height=40)
        self.btn_home.pack(side="left", padx=10)

    def on_review(self):
        if self.review_callback:
            self.review_callback()

    def show_results(self, total, s1, s2, s3, l_stats=None, s_stats=None, r_stats=None):
        self.lbl_total.configure(text=f"Total Score: {total}")
        self.lbl_listening.configure(text=f"Listening: {s1}")
        self.lbl_structure.configure(text=f"Structure: {s2}")
        self.lbl_reading.configure(text=f"Reading: {s3}")
        
        if l_stats and s_stats and r_stats:
            stats_text = (f"Correct Answers:\n"
                          f"Listening: {l_stats[0]}/{l_stats[1]}  |  "
                          f"Structure: {s_stats[0]}/{s_stats[1]}  |  "
                          f"Reading: {r_stats[0]}/{r_stats[1]}")
            self.lbl_stats.configure(text=stats_text)
        else:
            self.lbl_stats.configure(text="")

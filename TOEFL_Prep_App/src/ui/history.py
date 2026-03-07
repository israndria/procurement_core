import customtkinter as ctk

class HistoryFrame(ctk.CTkFrame):
    def __init__(self, master, back_callback, history_manager):
        super().__init__(master)
        self.back_callback = back_callback
        self.history_manager = history_manager
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Header
        self.header = ctk.CTkFrame(self, height=50)
        self.header.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        self.btn_back = ctk.CTkButton(self.header, text="< Back", width=60, command=self.back_callback)
        self.btn_back.pack(side="left", padx=10)
        
        self.lbl_title = ctk.CTkLabel(self.header, text="Test History", font=ctk.CTkFont(size=20, weight="bold"))
        self.lbl_title.pack(side="left", padx=20)

        # List Area
        self.scroll_frame = ctk.CTkScrollableFrame(self)
        self.scroll_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)

    def refresh(self):
        # Clear existing
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
            
        data = self.history_manager.load_history()
        # Sort by date desc
        data.reverse()
        
        if not data:
            ctk.CTkLabel(self.scroll_frame, text="No history yet.").pack(pady=20)
            return

        for i, entry in enumerate(data):
            # Create a card for each entry
            card = ctk.CTkFrame(self.scroll_frame)
            card.pack(fill="x", pady=5, padx=5)
            
            # Date
            date_str = entry.get("timestamp", "").split("T")[0]
            ctk.CTkLabel(card, text=date_str, width=100).pack(side="left", padx=10)
            
            # Score
            score = entry.get("total_score", 0)
            ctk.CTkLabel(card, text=f"Score: {score}", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=10)
            
            # Details
            s = entry.get("section_scores", {})
            details = f"L:{s.get('listening',0)} S:{s.get('structure',0)} R:{s.get('reading',0)}"
            ctk.CTkLabel(card, text=details).pack(side="left", padx=20)

import customtkinter as ctk

class DashboardFrame(ctk.CTkFrame):
    def __init__(self, master, start_callback, history_callback):
        super().__init__(master)
        
        self.start_callback = start_callback
        self.history_callback = history_callback
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        self.title = ctk.CTkLabel(self, text="TOEFL ITP Simulator", font=ctk.CTkFont(size=32, weight="bold"))
        self.title.grid(row=1, column=0, pady=20)

        self.btn_start = ctk.CTkButton(self, text="Start New Simulation", command=self.start_callback, height=50, font=ctk.CTkFont(size=18))
        self.btn_start.grid(row=2, column=0, pady=20)
        
        self.btn_history = ctk.CTkButton(self, text="View History", command=self.history_callback, height=40, font=ctk.CTkFont(size=16), fg_color="transparent", border_width=2)
        self.btn_history.grid(row=3, column=0, pady=10)

    def update_stats(self, best_score):
        # Could add a label here
        pass

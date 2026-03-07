import customtkinter as ctk
import os
import sys

# Ensure src is in path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from logic.scoring import ScoreCalculator
from logic.user_manager import HistoryManager
from ui.dashboard import DashboardFrame
from ui.simulation import SimulationFrame
from ui.result import ResultFrame
from ui.history import HistoryFrame
from ui.review import ReviewFrame

class TOEFLApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("TOEFL ITP Simulator")
        self.geometry("1000x800")
        
        # Set theme
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        # Initialize Logic
        self.scorer = ScoreCalculator()
        self.history_manager = HistoryManager()

        # Layout Container
        self.container = ctk.CTkFrame(self)
        self.container.pack(fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        self.active_questions = [] 
        self.last_answers = {} # Store answers for review
        
        # Initialize Frames
        self.frames["Dashboard"] = DashboardFrame(self.container, self.start_simulation, self.show_history)
        self.frames["Simulation"] = SimulationFrame(self.container, self.finish_simulation)
        self.frames["Review"] = ReviewFrame(self.container, self.show_result_screen) # Back to Result or Dashboard? Let's say Result
        self.frames["Result"] = ResultFrame(self.container, self.show_dashboard, self.start_review_mode)
        self.frames["History"] = HistoryFrame(self.container, self.show_dashboard, self.history_manager)

        for F in self.frames.values():
            F.grid(row=0, column=0, sticky="nsew")

        self.show_dashboard()

    def show_dashboard(self):
        self.frames["Dashboard"].tkraise()
        
    def show_result_screen(self):
        self.frames["Result"].tkraise()

    def start_simulation(self):
        frame = self.frames["Simulation"]
        # Load questions
        data_path = os.path.join(os.path.dirname(__file__), 'data', 'questions.json')
        frame.load_questions(data_path)
        self.active_questions = frame.questions 
        self.last_answers = {}
        frame.tkraise()

    def start_review_mode(self):
        frame = self.frames["Review"]
        frame.load_data(self.active_questions, self.last_answers)
        frame.tkraise()

    def finish_simulation(self, answers):
        self.last_answers = answers # Save for review
        
        # Grading Logic
        listening_correct = 0
        structure_correct = 0
        reading_correct = 0
        mistakes = []

        for q in self.active_questions:
            q_id = q.get("id")
            user_idx = answers.get(q_id, -1)
            
            # Determine correct index
            options = q.get("options", [])
            correct_ans_str = q.get("answer")
            correct_idx = -1
            try:
                correct_idx = options.index(correct_ans_str)
            except ValueError:
                pass # Answer string not found exactly?
            
            if user_idx == correct_idx:
                # Correct!
                # Basic section detection based on ID or file structure
                # My sample data has "section_2_structure" etc but here strict types might be better
                # Heuristic: Check ID prefix or assume structure since that's what we loaded mostly
                if str(q_id).startswith("L"):
                    listening_correct += 1
                elif str(q_id).startswith("S"):
                    structure_correct += 1
                elif str(q_id).startswith("R"):
                    reading_correct += 1
                else:
                    # Default/Fallback
                    structure_correct += 1
            else:
                mistakes.append(q_id)

        # Calculate Score
        total, s1, s2, s3 = self.scorer.calculate_score(listening_correct, structure_correct, reading_correct)
        
        # Save History
        self.history_manager.save_session("Simulation", total, (s1, s2, s3), mistakes)
        
        # Show Results with Raw Counts
        # We need to know total questions per section to show "Correct / Total"
        # Heuristic: S1=50, S2=40, S3=50 (approx, or calculate from active_questions)
        
        s1_total = sum(1 for q in self.active_questions if str(q.get("id")).startswith("L") or q.get("type") == "Listening")
        s2_total = sum(1 for q in self.active_questions if str(q.get("id")).startswith("S") or q.get("type") == "Structure")
        s3_total = sum(1 for q in self.active_questions if str(q.get("id")).startswith("R") or q.get("type") == "Reading")
        
        # Fallback if types aren't perfect
        if s1_total + s2_total + s3_total == 0:
            s2_total = len(self.active_questions) # Assume all structure if detection fails
            
        result_frame = self.frames["Result"]
        result_frame.show_results(total, s1, s2, s3, 
                                  (listening_correct, s1_total),
                                  (structure_correct, s2_total),
                                  (reading_correct, s3_total))
        result_frame.tkraise()

    def show_history(self):
        frame = self.frames["History"]
        frame.refresh()
        frame.tkraise()


if __name__ == "__main__":
    app = TOEFLApp()
    app.mainloop()

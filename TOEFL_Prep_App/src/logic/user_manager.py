import json
import os
from datetime import datetime

class HistoryManager:
    def __init__(self, data_path=None):
        if data_path is None:
            base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            data_path = os.path.join(base_path, 'data', 'history.json')
        
        self.history_file = data_path
        self._ensure_file_exists()

    def _ensure_file_exists(self):
        if not os.path.exists(self.history_file):
            with open(self.history_file, 'w') as f:
                json.dump([], f)

    def save_session(self, mode, total_score, section_scores, mistakes):
        """
        Saves a test session to history.
        section_scores: tuple or list (listening, structure, reading)
        mistakes: list of question IDs or details
        """
        entry = {
            "timestamp": datetime.now().isoformat(),
            "mode": mode,
            "total_score": total_score,
            "section_scores": {
                "listening": section_scores[0],
                "structure": section_scores[1],
                "reading": section_scores[2]
            },
            "mistakes": mistakes
        }

        history = self.load_history()
        history.append(entry)

        with open(self.history_file, 'w') as f:
            json.dump(history, f, indent=4)

    def load_history(self):
        try:
            with open(self.history_file, 'r') as f:
                return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError):
            return []

    def get_best_score(self):
        history = self.load_history()
        if not history:
            return 0
        return max(item.get('total_score', 0) for item in history)

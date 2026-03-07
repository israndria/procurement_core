import json
import os

class ScoreCalculator:
    def __init__(self, data_path=None):
        if data_path is None:
            # Default to relative path from this file
            base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            data_path = os.path.join(base_path, 'data', 'scoring_table.json')
        
        self.conversion_table = self._load_table(data_path)

    def _load_table(self, path):
        try:
            with open(path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"Error: Scoring table not found at {path}")
            return {}

    def calculate_score(self, listening_raw, structure_raw, reading_raw):
        """
        Calculates the TOEFL ITP total score.
        Formula: (Converted Score 1 + Converted Score 2 + Converted Score 3) * 10 / 3
        """
        s1 = self._convert_score('section_1_listening', listening_raw)
        s2 = self._convert_score('section_2_structure', structure_raw)
        s3 = self._convert_score('section_3_reading', reading_raw)

        total = (s1 + s2 + s3) * 10 / 3
        return int(round(total)), s1, s2, s3

    def _convert_score(self, section, raw_score):
        # The table keys are strings, so convert raw_score to string
        # If raw score is not in table (e.g. < 0 or > max), handle gracefully
        # Data keys are "50", "49", etc.
        
        table = self.conversion_table.get(section, {})
        
        # Clamp score to available keys just in case
        # Max scores: L=50, S=40, R=50. Min=0.
        # But JSON keys are strings. 
        
        # Simple lookup
        score_str = str(raw_score)
        if score_str in table:
            return table[score_str]
        
        # Fallback if specific score not found (shouldn't happen with full table)
        # Try to find closest or return 0
        return 20 # Minimum scaled score roughly

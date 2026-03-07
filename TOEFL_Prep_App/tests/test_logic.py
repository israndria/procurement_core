import sys
import os
import unittest

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'src'))

from logic.scoring import ScoreCalculator
from logic.user_manager import HistoryManager

class TestTOEFL(unittest.TestCase):
    def test_scoring(self):
        scorer = ScoreCalculator()
        # Test max score
        total, s1, s2, s3 = scorer.calculate_score(50, 40, 50)
        print(f"Max Score: {total} (L:{s1}, S:{s2}, R:{s3})")
        self.assertTrue(660 <= total <= 677)

        # Test min score (0 correct)
        total_min, m1, m2, m3 = scorer.calculate_score(0, 0, 0)
        print(f"Min Score: {total_min} (L:{m1}, S:{m2}, R:{m3})")
        self.assertTrue(total_min >= 210)

    def test_history(self):
        hm = HistoryManager()
        # Save a dummy session
        hm.save_session("Test Mode", 500, (50, 50, 50), ["S1"])
        
        # Load and verify
        history = hm.load_history()
        self.assertTrue(len(history) > 0)
        self.assertEqual(history[-1]['total_score'], 500)
        print("History save/load working.")

if __name__ == '__main__':
    unittest.main()

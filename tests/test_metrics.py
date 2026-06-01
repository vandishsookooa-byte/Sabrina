import unittest

import pandas as pd

from app import apply_filters, calculate_metrics, prepare_dataframe


class MetricsTestCase(unittest.TestCase):
    def setUp(self):
        self.df = pd.DataFrame(
            {
                "Date": ["2025-01-01", "2025-01-02", "2025-01-03", "2025-02-02"],
                "Department": ["Ops", "Ops", "HR", "Ops"],
                "Employee": ["A", "A", "B", "C"],
                "Attendance": ["Present", "Absent", "Present", "Present"],
                "OT1": [2, 0, 1, 1],
                "OT2": [0, 0, 0, 2],
                "OT3": [0, 1, 0, 0],
                "Leave Type": ["Annual", "Sick", "Annual", "Sick"],
                "Leave Quantity": [1, 0, 0, 0],
                "Missing Hours": [0.5, 1.0, 0.0, 0.2],
            }
        )

    def test_calculate_metrics_core_outputs(self):
        prepared, colmap = prepare_dataframe(self.df)
        metrics = calculate_metrics(prepared, colmap)

        self.assertIn("monthly", metrics)
        self.assertIn("rankings", metrics)
        self.assertGreater(len(metrics["monthly"]), 0)
        self.assertGreaterEqual(metrics["overall"]["attendance"], 0)
        self.assertLessEqual(metrics["overall"]["attendance"], 100)

    def test_filtering_by_department(self):
        prepared, colmap = prepare_dataframe(self.df)
        filtered = apply_filters(prepared, colmap, {"department": "HR"})
        self.assertEqual(set(filtered["Department"].unique()), {"HR"})


if __name__ == "__main__":
    unittest.main()

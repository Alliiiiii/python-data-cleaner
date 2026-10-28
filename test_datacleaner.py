"""Run with: python -m unittest test_datacleaner -v"""
import unittest
from datetime import date
from types import SimpleNamespace
from pathlib import Path
import datacleaner as dc

OPTS = SimpleNamespace(sheet=None, keys=None, phone_prefix="+92", dayfirst=True)
SAMPLE = Path(__file__).parent / "sample" / "messy_customers.csv"


class Parsers(unittest.TestCase):
    def test_numbers(self):
        self.assertEqual(dc.parse_number("Rs. 25,750.50"), 25750.5)
        self.assertEqual(dc.parse_number("(2,500)"), -2500)
        self.assertIsNone(dc.parse_number("abc"))

    def test_dates(self):
        self.assertEqual(dc.parse_date("13/05/2025", True), date(2025, 5, 13))
        self.assertEqual(dc.parse_date("2025-03-04", True), date(2025, 3, 4))
        self.assertIsNone(dc.parse_date("31/02/2025", True))
        self.assertIsNone(dc.parse_date("not sure", True))

    def test_phone_and_email(self):
        self.assertEqual(dc.fmt_phone("0300-1234567", "+92")[0], "+923001234567")
        self.assertIsNotNone(dc.fmt_email("bad.email.com")[1])
        self.assertEqual(dc.fmt_email(" A@B.com ")[0], "a@b.com")


class EndToEnd(unittest.TestCase):
    def test_sample_file(self):
        res = dc.clean_file(SAMPLE, OPTS)
        self.assertEqual(len(res["clean"]), 15)
        self.assertEqual(len(res["removed"]), 7)
        self.assertEqual(res["n_source"] - res["n_blank"] - len(res["removed"]), len(res["clean"]))
        self.assertEqual(res["filled"], 1)
        self.assertEqual(len(res["issues"]), 2)
        self.assertFalse(res["clean"]["email"].str.contains("[A-Z ]", regex=True).any())


if __name__ == "__main__":
    unittest.main()

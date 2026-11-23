import contextlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from screening import COLUMNS, main, validate_results


class ScreeningTests(unittest.TestCase):
    def result(self, row):
        return {
            "row": row,
            "fit": "Unknown",
            "recommendation": "MORE INFO",
            "reasoning": "The application is incomplete. Thesis fit is unclear.",
            "follow_up": "Where is the company headquartered?",
        }

    def test_rejects_missing_duplicate_and_invalid_results(self):
        rows = [{"row": 2}]
        result = self.result(2)
        for results in ([], [result, result], [self.result(3)], [dict(result, fit="Strong")]):
            with self.subTest(results=results), self.assertRaises(ValueError):
                validate_results(results, rows)

    def test_extract_and_export_keep_rows_links_and_source(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            source = folder / "inbound.xlsx"
            rows_file = folder / "rows.json"
            results = folder / "results.json"
            output = folder / "screened.xlsx"
            workbook = Workbook()
            workbook.active.title = "Notes"
            sheet = workbook.create_sheet("Applications")
            sheet.append(["Company Name", "Deck"])
            sheet.append(["First", "Pitch"])
            sheet["B2"].hyperlink = "https://example.com/deck.pdf"
            sheet["B4"] = "partial application"
            workbook.save(source)
            workbook.close()
            original = source.read_bytes()

            def run(*args):
                with patch.object(sys, "argv", ["screening.py", *map(str, args), "--sheet", "Applications"]):
                    with contextlib.redirect_stdout(io.StringIO()):
                        main()

            run("extract", source, rows_file)
            rows = json.loads(rows_file.read_text(encoding="utf-8"))
            self.assertEqual([row["row"] for row in rows], [2, 4])
            self.assertEqual(rows[0]["application"]["Deck"]["link"], "https://example.com/deck.pdf")
            results.write_text(json.dumps([self.result(4), self.result(2)]), encoding="utf-8")
            run("export", source, results, output)
            workbook = load_workbook(output)
            try:
                sheet = workbook["Applications"]
                self.assertEqual(workbook.sheetnames, ["Notes", "Applications"])
                self.assertEqual(sheet["B2"].hyperlink.target, "https://example.com/deck.pdf")
                self.assertEqual([cell.value for cell in sheet[1]][2:], list(COLUMNS.values()))
                self.assertEqual(sheet["D4"].value, "MORE INFO")
                self.assertIsNone(sheet["D3"].value)
                self.assertEqual(source.read_bytes(), original)
            finally:
                workbook.close()

import io
import json
import unittest
from unittest.mock import patch

from purdue_format_check import Issue, main, serialize_findings


class TestJsonReport(unittest.TestCase):
    def test_serialize_findings_empty_passes(self):
        """When no issues exist, serialize_findings returns PASS and 0 errors/infos."""
        result = serialize_findings([], "thesis.docx", "purdue")
        self.assertEqual(result["filename"], "thesis.docx")
        self.assertEqual(result["university"], "purdue")
        self.assertEqual(result["verdict"], "PASS")
        self.assertEqual(result["errors"], 0)
        self.assertEqual(result["infos"], 0)
        self.assertEqual(result["ada"], [])
        self.assertEqual(result["formatting"], [])

    def test_serialize_findings_structure_and_categorization(self):
        """serialize_findings groups issues, separates A11Y into ada, and computes verdict."""
        issues = [
            Issue(rule="H1.8.FONT", loc="p. 3", msg="Font is Arial, expected Times New Roman", severity="ERROR"),
            Issue(rule="H1.8.FONT", loc="p. 5", msg="Font is Calibri, expected Times New Roman", severity="ERROR"),
            Issue(rule="H1.10.INDENT", loc="p. 10", msg="Check indent", severity="INFO"),
            Issue(rule="A11Y.ALT", loc="fig. 1", msg="Image is missing alt text", severity="ERROR"),
        ]
        result = serialize_findings(issues, "my_thesis.docx", "purdue")

        self.assertEqual(result["filename"], "my_thesis.docx")
        self.assertEqual(result["university"], "purdue")
        self.assertEqual(result["verdict"], "FAIL")
        self.assertEqual(result["errors"], 3)
        self.assertEqual(result["infos"], 1)

        # Check ada vs formatting split
        ada_rules = [g["rule"] for g in result["ada"]]
        formatting_rules = [g["rule"] for g in result["formatting"]]
        self.assertEqual(ada_rules, ["A11Y.ALT"])
        self.assertEqual(formatting_rules, ["H1.8.FONT", "H1.10.INDENT"])

        # Check grouped items
        font_group = next(g for g in result["formatting"] if g["rule"] == "H1.8.FONT")
        self.assertEqual(font_group["severity"], "ERROR")
        self.assertEqual(len(font_group["items"]), 2)
        self.assertEqual(font_group["items"][0], {"loc": "p. 3", "msg": "Font is Arial, expected Times New Roman"})
        self.assertEqual(font_group["items"][1], {"loc": "p. 5", "msg": "Font is Calibri, expected Times New Roman"})

    def test_cli_json_flag_pass_exit_code_and_json_validity(self):
        """When check finds only INFO notes (0 errors), --json outputs valid JSON and exits 0."""
        sample_issues = [
            Issue(rule="H1.10.INDENT", loc="p. 4", msg="Manual review suggested", severity="INFO"),
        ]
        stdout_capture = io.StringIO()
        with patch("purdue_format_check.check", return_value=sample_issues), \
             patch("sys.stdout", stdout_capture):
            code = main(["purdue_format_check.py", "test.docx", "--json"])

        self.assertEqual(code, 0)
        output = stdout_capture.getvalue()
        parsed = json.loads(output)
        self.assertEqual(parsed["verdict"], "PASS")
        self.assertEqual(parsed["errors"], 0)
        self.assertEqual(parsed["infos"], 1)
        self.assertEqual(parsed["filename"], "test.docx")

    def test_cli_json_flag_fail_exit_code_on_errors(self):
        """When check finds ERROR issues, --json outputs valid JSON and exits nonzero (1)."""
        sample_issues = [
            Issue(rule="H1.8.FONT", loc="p. 1", msg="Wrong font", severity="ERROR"),
        ]
        stdout_capture = io.StringIO()
        with patch("purdue_format_check.check", return_value=sample_issues), \
             patch("sys.stdout", stdout_capture):
            code = main(["purdue_format_check.py", "thesis.docx", "--json", "--university", "purdue"])

        self.assertEqual(code, 1)
        output = stdout_capture.getvalue()
        parsed = json.loads(output)
        self.assertEqual(parsed["verdict"], "FAIL")
        self.assertEqual(parsed["errors"], 1)
        self.assertEqual(parsed["university"], "purdue")


if __name__ == "__main__":
    unittest.main()

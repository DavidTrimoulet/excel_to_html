from unittest import TestCase
import excel_to_html
from pathlib import Path
import xlrd


class TestGet_excel_to_html(TestCase):
    def setUp(self) :
        self.p = Path('.')
        self.path_to_input = self.p / "Catalogue Formations Roboticipation.xlsx"
        self.path_to_output = self.p / ".." / "test1.html"

    def test_get_input_output_file(self):
        input, output = excel_to_html.get_input_output_file(["excel_to_html.py", "Catalogue Formations Roboticipation.xlsx", "../test1.html"])
        self.assertEqual(input, self.path_to_input)
        self.assertEqual(output, self.path_to_output)

    def test_convert_xls_to_html(self):
        print(self.path_to_input)
        xls_file = xlrd.open_workbook(self.path_to_input)
        html = excel_to_html.convert_xls_to_html(xls_file)
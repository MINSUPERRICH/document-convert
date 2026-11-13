import ast
import io
import re
import unittest
from decimal import Decimal
from pathlib import Path

import pandas as pd
from docx import Document
from openpyxl import Workbook, load_workbook

# Load only the new pure helpers: importing app.py runs the Streamlit UI.
source = Path(__file__).resolve().parents[1].joinpath('app.py').read_text(encoding='utf-8')
parsed = ast.parse(source)
nodes = [n for n in parsed.body if isinstance(n, ast.FunctionDef) and n.name.startswith('packing_')]
namespace = {'pd': pd, 'io': io, 're': re, 'Decimal': Decimal, 'Document': Document}
exec(compile(ast.Module(body=nodes, type_ignores=[]), 'app.py', 'exec'), namespace)
extract, total, export = (namespace['packing_extract'], namespace['packing_totals'], namespace['packing_export'])


class PackingTests(unittest.TestCase):
    def setUp(self):
        self.mapping = {'Pieces': 'Column 1', 'Weight': 'Column 2', 'Amount': 'Column 3'}
        self.units = {'Pieces': ('fixed', 'pcs'), 'Weight': ('column', 'Column 4'), 'Amount': ('fixed', 'USD')}
        self.frame = pd.DataFrame([
            {'Group': 'A', 'Column 1': '2', 'Column 2': '0.1', 'Column 3': '1.20', 'Column 4': 'kg'},
            {'Group': 'A', 'Column 1': '3', 'Column 2': '0.2', 'Column 3': '2.30', 'Column 4': 'kg'},
        ])

    def test_exact_totals_and_validation(self):
        result, errors = total(self.frame, self.mapping, self.units)
        self.assertEqual(errors, [])
        self.assertEqual(list(result.loc[0, ['Pieces', 'Weight', 'Amount']]), ['5', '0.3', '3.50'])
        for col, value in [('Group', ''), ('Group', float('nan')), ('Column 1', ''), ('Column 2', 'NaN'), ('Column 3', '1,000'), ('Column 4', '')]:
            bad = self.frame.copy()
            bad.loc[0, col] = value
            result, errors = total(bad, self.mapping, self.units)
            self.assertTrue(errors, (col, value))
            self.assertTrue(result.empty)
        mixed = self.frame.copy()
        mixed.loc[1, 'Column 4'] = 'lb'
        self.assertIn('incompatible', ' '.join(total(mixed, self.mapping, self.units)[1]))
        mixed.loc[1, 'Group'] = 'B'
        self.assertFalse(total(mixed, self.mapping, self.units)[1])

    def test_docx_and_xlsx_extraction(self):
        doc = Document()
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = 'header'
        table.cell(1, 1).text = 'value'
        data = io.BytesIO()
        doc.save(data)
        frame = extract(data.getvalue(), 'docx')
        self.assertEqual(frame.loc[1, 'Column 2'], 'value')
        self.assertEqual(frame.loc[1, 'Source row'], '2')
        book = Workbook()
        book.active.title = 'First'
        book.active.append(['one', '=1+1'])
        book.create_sheet('Second').append(['two'])
        data = io.BytesIO()
        book.save(data)
        frame = extract(data.getvalue(), 'xlsx', ['First'])
        self.assertEqual(frame.loc[0, 'Column 2'], '=1+1')
        self.assertEqual(frame.loc[0, 'Source'], 'First')
        self.assertTrue(extract(data.getvalue(), 'xlsx', []).empty)

    def test_export_formula_injection(self):
        data = pd.DataFrame({'=header': ['=1+1', '  +cmd', '\t@cmd', '-1', 'normal']})
        book = load_workbook(io.BytesIO(export(data, pd.DataFrame({'Group': ['=x']}))), data_only=False)
        for sheet in book:
            for row in sheet:
                for cell in row:
                    self.assertNotEqual(cell.data_type, 'f')
        self.assertEqual(book['Rows']['A2'].value, "'=1+1")
        self.assertEqual(book['Totals']['A2'].value, "'=x")

    def test_pdf_pages(self):
        try:
            from reportlab.pdfgen import canvas
            import pdfplumber  # noqa: F401
        except ImportError:
            self.skipTest('pdfplumber and reportlab required for PDF fixture')
        data = io.BytesIO()
        pdf = canvas.Canvas(data, pagesize=(200, 200))
        for page in (1, 2):
            for x in (20, 100, 180):
                pdf.line(x, 100, x, 180)
            for y in (100, 140, 180):
                pdf.line(20, y, 180, y)
            pdf.drawString(30, 155, 'H')
            pdf.drawString(30, 115, str(page))
            pdf.showPage()
        pdf.save()
        all_rows = extract(data.getvalue(), 'pdf')
        self.assertEqual(set(all_rows['Source']), {'Page 1', 'Page 2'})
        only_two = extract(data.getvalue(), 'pdf', [2])
        self.assertEqual(set(only_two['Source']), {'Page 2'})
        self.assertTrue(extract(data.getvalue(), 'pdf', []).empty)
        from PIL import ImageChops
        preview = namespace['packing_pdf_preview']
        first = preview(data.getvalue(), 1)
        second = preview(data.getvalue(), 2)
        self.assertEqual(first.size, (300, 300))
        self.assertIsNotNone(ImageChops.difference(first, second).getbbox())
        with self.assertRaises(ValueError):
            preview(data.getvalue(), 0)


if __name__ == '__main__':
    unittest.main()

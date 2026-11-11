"""离线回归检查：python -m unittest discover -s tests -v"""
import ast
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
if not (ROOT / 'generate_valuation_sheet.py').exists():
    ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
import generate_valuation_sheet as gen


class WorkflowTests(unittest.TestCase):
    def test_expansion_preserves_values_and_moves_references(self):
        with tempfile.TemporaryDirectory() as d:
            main, source = Path(d) / 'main.xlsx', Path(d) / 'input.xlsx'
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = '测试页'
            gen.generate_sheet(ws, [{'name': '苹果', 'code': 'AAPL', 'weight': 1}])
            ws['D5'], ws['E5'] = 20, 15
            overview = wb.create_sheet('总览对比')
            overview['B45'] = "='测试页'!$B$85"
            overview['C1'] = '="测试页!B85"'
            overview['D1'] = "='测试页'!B5"
            history = wb.create_sheet('当一家公司历史')
            history['A5'] = '旧快照'
            wb.save(main)
            src = openpyxl.Workbook()
            src.active.title = '测试页'
            src.active.append(['名称', '代码', '净值权重'])
            for n in range(100):
                src.active.append(['公司' + str(n), 'AAPL' if n == 0 else 'T' + str(n), .01])
            src.save(source)
            gen.merge_into_main(source, main)
            result = openpyxl.load_workbook(main)
            self.assertEqual(result['总览对比']['B45'].value, "='测试页'!$B$132")
            self.assertEqual(result['总览对比']['C1'].value, '="测试页!B85"')
            self.assertEqual(result['总览对比']['D1'].value, "='测试页'!B5")
            self.assertEqual(result['测试页']['D5'].value, 20)
            self.assertEqual(result['测试页']['A150'].value, '当一家公司历史')
            self.assertEqual(result['当一家公司历史']['A5'].value, '旧快照')
            self.assertEqual(result.sheetnames, wb.sheetnames)

    def test_mapping_path_is_script_relative(self):
        self.assertEqual(Path(gen.GF_MAP_PATH), ROOT / 'gf_ticker_map.csv')
        self.assertGreater(len(gen.GF_RULES), 100)

    def test_negative_ev_and_zero_boundary(self):
        # Load pure functions without starting browsers or requiring Playwright.
        tree = ast.parse((ROOT / 'update_gf_valuations.py').read_text(encoding='utf-8'))
        names = {'include_flag', 'iter_holding_rows', '_as_one_ratio'}
        module = ast.Module(body=[node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in names], type_ignores=[])
        ns = {}
        exec(compile(module, '<pure valuation functions>', 'exec'), ns)
        self.assertEqual(ns['include_flag'](-20, 10, False, 'ev'), 1)
        self.assertEqual(ns['include_flag'](0, 10, False, 'ev'), 0)
        self.assertEqual(ns['include_flag'](-20, 10, True, 'ev'), 0)
        ws = openpyxl.Workbook().active
        for row, weight, current in [(5, .8, 10), (6, .2, -20)]:
            for col, value in [(1, '股票'), (2, 'T' + str(row)), (3, weight), (16, current), (17, 10), (19, 1)]:
                ws.cell(row, col, value)
        self.assertAlmostEqual(ns['_as_one_ratio'](ws, 16, 17, 19, True), .1 / .07)
        ws['P5'] = -10
        self.assertIsNone(ns['_as_one_ratio'](ws, 16, 17, 19, True))


if __name__ == '__main__':
    unittest.main()

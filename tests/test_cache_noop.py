import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from datetime import date
import openpyxl

ROOT = Path(__file__).resolve().parent.parent
if not (ROOT / 'update_gf_valuations.py').exists():
    ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
import update_gf_valuations as updater


class CacheNoOpTests(unittest.TestCase):
    def setUp(self):
        self.wb = openpyxl.Workbook()
        self.ws = self.wb.active
        self.ws.title = '组合'
        self.ws['A5'], self.ws['B5'], self.ws['C5'] = 'Apple', 'AAPL', 1
        self.cache = {}
        for field, cur, med, flag, slugs in updater.TERM_SPECS:
            self.ws.cell(5, cur, 20)
            self.ws.cell(5, med, 10)
            self.ws.cell(5, flag, 1)
            self.cache['AAPL|' + slugs[0]] = {
                'current': 20, 'med': 10, 'fetched_at': date.today().isoformat()}
        self.ws['T5'] = 'GF自动 2026-11-11；旧备注'

    def write(self, path, snapshot=True, dry_run=False):
        updater.maybe_write_ready_sheets(self.wb, path, ['组合'], {}, self.cache, .01, dry_run, snapshot)

    def test_same_cache_does_not_save_date_or_history(self):
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / 'book.xlsx'
            self.wb.save(path)
            before = path.read_bytes()
            with patch.object(updater, 'snapshot_as_one_company') as history, patch.object(updater, 'refresh_overview_rank') as rank:
                self.write(path)
                history.assert_not_called()
                rank.assert_not_called()
            self.assertEqual(path.read_bytes(), before)
            self.assertEqual(self.ws['T5'].value, 'GF自动 2026-11-11；旧备注')

    def test_intermediate_change_snapshots_once_at_end(self):
        self.cache['AAPL|pettm']['current'] = 21
        with tempfile.TemporaryDirectory() as d, patch.object(updater, 'snapshot_as_one_company') as history, patch.object(updater, 'refresh_overview_rank'):
            path = Path(d) / 'book.xlsx'
            self.write(path, snapshot=False)
            history.assert_not_called()
            self.assertEqual(self.ws['D5'].value, 21)
            self.write(path)
            history.assert_called_once()
            self.write(path)
            history.assert_called_once()

    def test_missing_values_and_flag_changes_are_updates(self):
        self.ws['D5'] = None
        self.ws['S5'] = 0
        self.assertEqual(updater.flush_cache_to_sheet(self.ws, {}, self.cache), 1)
        self.assertEqual(self.ws['D5'].value, 20)
        self.assertEqual(self.ws['S5'].value, 1)
        self.assertEqual(updater.flush_cache_to_sheet(self.ws, {}, self.cache), 0)

    def test_dry_run_is_unchanged(self):
        self.cache['AAPL|pettm']['current'] = 21
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / 'book.xlsx'
            self.write(path, dry_run=True)
            self.assertFalse(path.exists())
            self.assertEqual(self.ws['D5'].value, 20)
            self.assertEqual(self.ws['T5'].value, 'GF自动 2026-11-11；旧备注')

    def test_process_sheet_same_values_not_counted_as_write(self):
        class Client:
            blocked = False
            stalled = False
            def fetch_term(self, *args, **kwargs):
                return {'current': 20, 'med': 10}
        stats = {'fetch': 0, 'ok': 0, 'miss': 0, 'wrote': 0, 'skip': 0}
        updater.process_sheet(self.ws, Client(), {}, .01, False, 0, stats)
        self.assertEqual(stats['wrote'], 0)
        self.assertEqual(self.ws['T5'].value, 'GF自动 2026-11-11；旧备注')

if __name__ == '__main__':
    unittest.main()

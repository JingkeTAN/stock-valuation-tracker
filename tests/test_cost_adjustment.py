import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import openpyxl
from openpyxl.chart import BarChart, Reference
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import generate_valuation_sheet as gen
import update_gf_valuations as upd
from valuation_costs import cost_factor, ensure_cost_parameters, cost_formulas


class CostTests(unittest.TestCase):
    def book(self):
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = '总览对比'
        for col, name in [('B', '低费'), ('C', '高费')]:
            wb.create_sheet(name)
            ws[f'{col}45'] = f"='{name}'!B85"
        for r, col, name in [(54,'B','低费'), (55,'C','高费')]:
            ws.cell(r,1,name)
            for c in range(2,6): ws.cell(r,c,f'={col}{43+c}')
        ws['A56']='横坐标已按四维平均'
        ensure_cost_parameters(wb,['低费','高费'])
        wb['基金参数']['C5']=0
        wb['基金参数']['C6']=.014
        return wb

    def test_factor_missing_duplicate_invalid_and_zero(self):
        wb=self.book(); ws=wb['基金参数']
        self.assertAlmostEqual(cost_factor(wb,'高费'),1.2)
        self.assertEqual(cost_factor(wb,'低费'),1)
        for bad in (None, '', -.01, float('nan'), '=1/100'):
            ws['C6']=bad
            self.assertIsNone(cost_factor(wb,'高费'))
        ws['C6']=.014
        for bad in (0, -.07, None, '7%'):
            ws['B2']=bad
            self.assertIsNone(cost_factor(wb,'高费'))
        ws['B2']=.07
        ws.append(['高费','',.01])
        self.assertIsNone(cost_factor(wb,'高费'))

    def test_cost_changes_order_and_formulas_follow_identity(self):
        wb=self.book(); ws=wb['总览对比']
        with patch.object(upd,'compute_as_one_pack',side_effect=lambda s:(1,1,1,1) if s.title=='低费' else (.9,.9,.9,.9)):
            upd.refresh_overview_rank(wb)
            self.assertEqual(ws['A54'].value,'低费')
            self.assertIn('低费',ws['I54'].value)
            self.assertIn('H54',ws['J54'].value)
            wb['基金参数']['C6']=0
            upd.refresh_overview_rank(wb)
            self.assertEqual(ws['A54'].value,'高费')
            self.assertIn('高费',ws['I54'].value)
            wb['基金参数']['C6']=None
            upd.refresh_overview_rank(wb)
            self.assertEqual(ws['A55'].value,'高费')

    def test_holdings_growth_preserves_parameters_and_history(self):
        with tempfile.TemporaryDirectory() as d:
            main, src=Path(d)/'main.xlsx',Path(d)/'input.xlsx'
            w=openpyxl.Workbook();ws=w.active;ws.title='越南'
            info=gen.generate_sheet(ws,[{'name':'股票','code':'T'+str(i),'weight':.1} for i in range(10)])
            gen.generate_overview(w,[('越南',info)])
            ensure_cost_parameters(w,['越南']);p=w['基金参数']
            p['B5']='008763';p['C5']=.0166;p['B2']=.06
            h=w.create_sheet('当一家公司历史');h['A5']='原历史';w.save(main)
            for count in (18,100,10):
                s=openpyxl.Workbook();s.active.title='越南';s.active.append(['名称','代码','净值权重'])
                for i in range(count):s.active.append(['股票','T'+str(i),1/count])
                s.save(src);gen.merge_into_main(src,main)
                out=openpyxl.load_workbook(main)
                self.assertEqual(out['基金参数']['C5'].value,.0166)
                self.assertEqual(out['基金参数']['B5'].value,'008763')
                self.assertEqual(out['基金参数']['B2'].value,.06)
                self.assertEqual(out['当一家公司历史']['A5'].value,'原历史')
                self.assertIn('越南',out['总览对比']['I54'].value)

    def test_cost_sheet_excluded_from_updates(self):
        self.assertIn('基金参数',upd.SKIP_SHEETS)

    def test_existing_chart_uses_adjusted_dimensions_and_stays_bound(self):
        w=self.book();s=w['总览对比']
        chart=BarChart();chart.title='原四维图';chart.style=12
        chart.add_data(Reference(s,min_col=2,max_col=5,min_row=53,max_row=55),titles_from_data=True)
        chart.set_categories(Reference(s,min_col=1,min_row=54,max_row=55))
        s.add_chart(chart,'A75')
        with patch.object(upd,'compute_as_one_pack',return_value=(1,1,1,1)):
            for _ in range(2):upd.refresh_overview_rank(w)
        self.assertIs(s._charts[0],chart)
        self.assertEqual(chart.style,12)
        self.assertEqual(chart.anchor,'A75')
        self.assertEqual([x.val.numRef.f for x in chart.series],
                         [f"'总览对比'!${col}$54:${col}$55" for col in 'KLMN'])
        self.assertIn('B54*(1+I54/',s['K54'].value)
        self.assertIn('E54*(1+I54/',s['N54'].value)
        self.assertIsNone(chart.y_axis.scaling.max)

    def test_new_chart_uses_adjusted_cache_and_preserves_missing_dimensions(self):
        w=self.book();s=w['总览对比']
        with patch.object(upd,'compute_as_one_pack',return_value=(1,1,1,None)):
            upd.refresh_overview_rank(w)
        bars=s._charts[0]._charts[1].series
        self.assertIn('$K$54',bars[0].val.numRef.f)
        self.assertEqual(len(s._charts),2)
        self.assertIn('$B$54',s._charts[1]._charts[1].series[0].val.numRef.f)
        self.assertIn('ISNUMBER(E54)',s['N54'].value)
        with tempfile.TemporaryDirectory() as d:
            path=Path(d)/'two.xlsx';w.save(path)
            again=openpyxl.load_workbook(path)
            with patch.object(upd,'compute_as_one_pack',return_value=(1,1,1,None)):
                upd.refresh_overview_rank(again)
                upd.refresh_overview_rank(again)
            self.assertEqual(len(again['总览对比']._charts),2)
            self.assertIn('$B$54',again['总览对比']._charts[1]._charts[1].series[0].val.numRef.f)

    def test_repair_stale_helper_and_overview_references(self):
        w=openpyxl.Workbook();s=w.active;s.title='测试'
        for r in (5,6):
            for c,v in [(1,'公司'),(2,'T'+str(r)),(3,.5),(4,20),(5,10),(7,1)]:s.cell(r,c,v)
        s['A20']='当一家公司：组合PE / 常态PE'
        s['AI6']='=1/D5'
        o=w.create_sheet('总览对比');o['B45']="='测试'!B21"
        upd.repair_ranking_sources(w)
        self.assertEqual(o['B45'].value,"='测试'!B20")
        self.assertIn('C5:C6',s['B20'].value)
        self.assertIn('1/D6',s['AI6'].value)
        self.assertIn('ISNUMBER(E6)',s['AI6'].value)


if __name__=='__main__': unittest.main()

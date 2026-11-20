import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from copy import copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import generate_valuation_sheet as gen
import update_gf_valuations as upd

class FormattingTests(unittest.TestCase):
    def test_holding_and_summary_styles_follow_identity(self):
        with tempfile.TemporaryDirectory() as d:
            main, src = Path(d)/'main.xlsx', Path(d)/'input.xlsx'
            wb=openpyxl.Workbook(); ws=wb.active; ws.title='测试'
            gen.generate_sheet(ws,[{'name':'苹果','code':'AAPL','weight':.6}, {'name':'微软','code':'MSFT','weight':.4}])
            ws['D5'].font=Font(name='宋体',size=16,bold=True,color='FF112233')
            ws['D5'].fill=PatternFill('solid',fgColor='FFFF0000')
            ws['D5'].border=Border(bottom=Side(style='thick',color='FF112233'))
            ws['D5'].alignment=Alignment(horizontal='right')
            ws['D5'].number_format='0.00000'
            saved_style=copy(ws['D5']._style)
            ws.row_dimensions[5].height=33
            ws.column_dimensions['D'].width=24;ws.column_dimensions['AA'].hidden=True
            ws['B83'].fill=PatternFill('solid',fgColor='FF00FF00')
            ws.freeze_panes='D5'
            ws.merge_cells('A1:D1')
            ws.conditional_formatting.add('D5:D6',CellIsRule(operator='greaterThan',formula=['10'],fill=PatternFill('solid',fgColor='FFEEEE00')))
            wb.save(main)
            inp=openpyxl.Workbook();inp.active.title='测试';inp.active.append(['名称','代码','净值权重'])
            # Move AAPL, add 98 new holdings, force summary relocation.
            inp.active.append(['微软','MSFT',.5]); inp.active.append(['苹果','AAPL',.4])
            for i in range(98):inp.active.append(['新增','T'+str(i),.001])
            inp.save(src);gen.merge_into_main(src,main)
            result=openpyxl.load_workbook(main);out=result['测试']
            self.assertEqual(out['D6']._style,saved_style)
            self.assertEqual(out.row_dimensions[6].height,33)
            self.assertEqual(out.column_dimensions['D'].width,24)
            self.assertTrue(out.column_dimensions['AA'].hidden)
            self.assertEqual(out.freeze_panes,'D5')
            self.assertIn('A1:D1',str(out.merged_cells))
            self.assertEqual(out['B130'].fill.fgColor.rgb,'FF00FF00')
            self.assertEqual(str(next(iter(out.conditional_formatting)).sqref),'D5:D104')
            self.assertEqual(out['D7'].number_format,'General')

    def test_update_changes_only_font_colour(self):
        ws=openpyxl.Workbook().active
        ws['D5'].font=Font(name='宋体',size=17,bold=True,italic=True,color='FFFF0000')
        ws['D5'].fill=PatternFill('solid',fgColor='FF00FF00')
        upd.set_blue(ws['D5'],42)
        self.assertEqual(ws['D5'].font.name,'宋体')
        self.assertEqual(ws['D5'].font.sz,17)
        self.assertTrue(ws['D5'].font.bold and ws['D5'].font.italic)
        self.assertEqual(ws['D5'].font.color.rgb[-6:],'0000FF')
        ws['T5'].font=Font(name='微软雅黑',size=15,color='FFCC0000')
        style=copy(ws['T5']._style)
        upd.set_row_note(ws,5,['GF自动 2026-11-20'])
        self.assertEqual(ws['T5']._style,style)

    def test_ranking_moves_styles_and_preserves_charts(self):
        wb=openpyxl.Workbook();ws=wb.active;ws.title='总览对比'
        wb.create_sheet('贵');wb.create_sheet('便宜')
        ws['B45']="='贵'!B85";ws['C45']="='便宜'!B85"
        ws['A54']='贵';ws['B54']='=B45';ws['A55']='便宜';ws['B55']='=C45'
        ws['A56']='横坐标已按四维平均'
        ws['A54'].fill=PatternFill('solid',fgColor='FFFF0000');ws.row_dimensions[54].height=31
        ws['B54'].number_format='0.0000'
        chart=BarChart();chart.title='我的自定义图';chart.style=12
        chart.add_data(Reference(ws,min_col=2,min_row=54,max_row=55))
        chart.set_categories(Reference(ws,min_col=1,min_row=54,max_row=55));ws.add_chart(chart,'J80')
        with patch.object(upd,'compute_as_one_pack',side_effect=lambda s: (2,2,2,2) if s.title=='贵' else (.5,.5,.5,.5)):
            upd.refresh_overview_rank(wb)
        self.assertEqual(ws['A55'].value,'贵')
        self.assertEqual(ws['A55'].fill.fgColor.rgb,'FFFF0000')
        self.assertEqual(ws.row_dimensions[55].height,31)
        self.assertEqual(ws['B55'].number_format,'0.0000')
        self.assertIs(ws._charts[0],chart)
        self.assertEqual(chart.style,12);self.assertEqual(chart.anchor,'J80')
        with tempfile.TemporaryDirectory() as d:
            path=Path(d)/'result.xlsx';wb.save(path)
            self.assertEqual(len(openpyxl.load_workbook(path)['总览对比']._charts),1)

if __name__=='__main__': unittest.main()

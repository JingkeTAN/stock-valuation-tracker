"""Workbook-owned cost inputs shared by both update scripts (rates are not hardcoded)."""
from copy import copy
import math
from openpyxl.styles import Font, PatternFill

COST_SHEET = "基金参数"


def valid_number(value, positive=False):
    return (isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(value) and (value > 0 if positive else value >= 0))


def ensure_cost_parameters(wb, names):
    if COST_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(COST_SHEET)
        ws.append(["基金成本参数"])
        ws.append(["统一预期年化（费用前）", 0.07])
        ws.append(["年化成本率含运作费率及分红税损耗；调整评分=原四维平均×(1+成本率/预期年化)。"])
        ws.append(["估值页名称", "对应ETF/基金代码", "年化成本率", "来源/备注"])
        ws['B2'].number_format = '0.00%'
        ws['B2'].font = Font(color='0000FF')
        for cell in ws[4]:
            cell.fill = PatternFill('solid', fgColor='4472C4')
            cell.font = Font(bold=True, color='FFFFFF')
        for col, width in [('A', 29), ('B', 25), ('C', 20), ('D', 55)]:
            ws.column_dimensions[col].width = width
        ws.freeze_panes = 'C5'
    ws = wb[COST_SHEET]
    existing = {ws.cell(r, 1).value for r in range(5, ws.max_row + 1)}
    for name in names:
        if name in existing:
            continue
        r = max(5, ws.max_row + 1)
        ws.cell(r, 1, name)
        ws.cell(r, 4, '待填写对应产品代码和年化成本率；不默认按零成本计算')
        ws.cell(r, 3).number_format = '0.0000%'
        ws.cell(r, 3).font = Font(color='0000FF')
        existing.add(name)


def cost_factor(wb, name):
    """None means missing/invalid/duplicate inputs, including formula-valued inputs."""
    if COST_SHEET not in wb.sheetnames:
        return None
    ws = wb[COST_SHEET]
    expected = ws['B2'].value
    matches = [ws.cell(r, 3).value for r in range(5, ws.max_row + 1)
               if ws.cell(r, 1).value == name]
    if not valid_number(expected, positive=True) or len(matches) != 1:
        return None
    cost = matches[0]
    return 1 + cost / expected if valid_number(cost) else None


def cost_formulas(name, row):
    key = '"' + name.replace('"', '""') + '"'
    # Whole columns allow the parameter table to grow and be sorted manually.
    match = f'MATCH({key},\'{COST_SHEET}\'!$A:$A,0)'
    value = f'INDEX(\'{COST_SHEET}\'!$C:$C,{match})'
    rate = f'=IFERROR(IF(AND(COUNTIF(\'{COST_SHEET}\'!$A:$A,{key})=1,COUNT({value})=1,ISNUMBER({value}),{value}>=0),{value},"待填写/检查参数"),"待填写/检查参数")'
    expected = f"'{COST_SHEET}'!$B$2"
    score = f'=IFERROR(IF(AND(ISNUMBER(H{row}),ISNUMBER(I{row}),ISNUMBER({expected}),{expected}>0),H{row}*(1+I{row}/{expected}),"待填写/检查参数"),"待填写/检查参数")'
    return rate, score


def write_cost_columns(wb, ws, row, name):
    for col, formula in zip((9, 10), cost_formulas(name, row)):
        cell = ws.cell(row, col)
        if not cell.has_style:
            cell._style = copy(ws.cell(row, 8)._style)
        cell.value = formula
        cell.number_format = '0.0000%' if col == 9 else '0.0000'


def write_cost_headers(ws, note_row):
    for col, title in ((9, '年化成本率'), (10, '成本调整评分')):
        cell = ws.cell(53, col)
        if not cell.has_style:
            cell._style = copy(ws.cell(53, 8)._style)
        cell.value = title
        if ws.column_dimensions[cell.column_letter].width < 19:
            ws.column_dimensions[cell.column_letter].width = 19
    ws['A52'] = '估值对比（按成本调整评分从低到高；四维及图表保留原始估值）'
    ws.cell(note_row, 1, '横坐标已按四维平均的成本调整评分排序；图表仍为原始估值。成本参数在「基金参数」页修改；手改参数后运行 --rank-only 刷新顺序。')

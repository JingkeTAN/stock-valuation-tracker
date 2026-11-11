#!/usr/bin/env python3
"""
估值表生成器：从「输入 xlsx」批量生成完整计算页 + 总览 + 辅助页
输入文件：每个 sheet 只需三列 —— 名称 / 代码 / 净值权重
输出：同名多 sheet 的完整计算表 + 总览对比 + 网页更新来源 + 口径说明 + 当一家公司历史总账
合并模式：更新=只重生计算页（不动总览）；添加=新建页+总览追加列
进标志优先读取 gf_ticker_map.csv
兼容 Python 3.6+
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import os
import sys
from datetime import datetime

# ========== 配置区 ==========

GF_MAP_PATH = "GF估值更新器/gf_ticker_map.csv"
SUMMARY_START = 60

BANK_INSURANCE_CODES = {
    "TCB", "STB", "MBB", "VPB", "ACB", "HDB", "VCB", "CTG",
    "GS", "JPM", "BAC", "WFC", "C", "MS", "AXP", "TRV",
    "SAN", "INGA", "UCG", "HDFC", "ICICI", "SBI", "AXIS",
}

SPECIAL_EXCLUSIONS = {
    "UNH": ["PB"],
    "SPCX": ["PE", "PB", "PS", "EV"],
}

HEADERS = [
    "名称", "代码", "净值权重",
    "PE现", "PE十年中位", "PE/Med", "进PE",
    "PB现", "PB十年中位", "PB/Med", "进PB",
    "PS现", "PS十年中位", "PS/Med", "进PS",
    "EV/FCF现", "EV/FCF十年中位", "EV/FCF/Med", "进EVFCF",
    "备注",
    "进调和PE", "反推ROE", "反推ROE中位", "ΔROE(ppt)", "进ΔROE",
    "进调和PB",
    "lnPEMed", "lnPBMed", "lnPSMed", "lnEVMed",
    "invPEMed", "invPBMed", "invPSMed", "invEVMed",
    "invPE", "invPEmed", "invPB", "invPBmed",
    "invPS", "invPSmed", "invEV", "invEVmed",
]

COL_ALIASES = {
    "name": ["名称", "name", "股票", "证券名称"],
    "code": ["代码", "code", "ticker", "证券代码"],
    "weight": ["净值权重", "权重", "weight", "份额", "占比"],
}

SOURCES = [
    ("159687_亚太精选", "stockanalysis.com SGX:LCU holdings", "https://stockanalysis.com/quote/sgx/LCU/holdings/"),
    ("520870_巴西IBOV", "etfsbrasil.com.br BOVV11 / IT Now 官方", "https://www.etfsbrasil.com.br/etfs/bovv11"),
    ("159329_沙特", "companiesmarketcap.com CSOP Saudi Arabia ETF", "https://companiesmarketcap.com/csop-saudi-arabia-etf/holdings/"),
    ("513800_南方TOPIX", "minkabu.jp One ETF TOPIX (1473)", "https://minkabu.jp/stock/1473/fund_selection"),
    ("03431_港韩科技", "companiesmarketcap.com CSOP FTSE HK-Korea Tech+", "https://companiesmarketcap.com/hkd/csop-ftse-hk-korea-tech-index-etf/holdings/"),
    ("513070_港股通消费", "中证指数官网 港股通消费 931454", "https://www.csindex.com.cn/zh-CN/indices/index-detail/931454"),
    ("009225_中美互联网", "中证指数官网 中美互联网 930794", "https://www.csindex.com.cn/zh-CN/indices/index-detail/930794"),
    ("162415_美国消费", "SSGA XLY 每日持仓", "https://www.ssga.com/library-content/products/fund-data/etfs/us/holdings-daily-us-en-xly.xlsx"),
    ("513290_纳指生物", "Invesco IBBQ / NBI", "https://www.invesco.com/us/en/financial-products/etfs/invesco-nasdaq-biotechnology-etf.html#Portfolio"),
    ("161128_标普信息科技", "iShares S&P 500 Info Tech UCITS", "https://www.ishares.com/uk/individual/en/products/280510/ishares-s-p-500-information-technology-sector-ucits-etf/latest-holdings.csv"),
    ("513850_美国50", "stockanalysis.com SHA:513850 holdings", "https://stockanalysis.com/quote/sha/513850/holdings/"),
]

# ========== 规则加载 ==========

def load_gf_map(path):
    rules = {}
    if not os.path.exists(path):
        return rules
    try:
        import csv
        with open(path, encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                code = str(row.get("local_code", "")).strip().upper()
                if not code or code.startswith("#"):
                    continue
                rules[code] = {
                    "no_pe": str(row.get("no_pe", "")).strip() == "1",
                    "no_pb": str(row.get("no_pb", "")).strip() == "1",
                    "no_ps": str(row.get("no_ps", "")).strip() == "1",
                    "no_ev": str(row.get("no_ev", "")).strip() == "1",
                    "skip": str(row.get("skip", "")).strip() == "1",
                }
    except Exception as e:
        print("读取 gf_ticker_map 失败:", e)
    return rules

GF_RULES = load_gf_map(GF_MAP_PATH)

def get_inclusion_flags(code):
    code_upper = str(code).upper().strip()
    pe, pb, ps, ev = 1, 1, 1, 1
    if code_upper in GF_RULES:
        r = GF_RULES[code_upper]
        if r.get("skip"):
            return 0, 0, 0, 0
        if r["no_pe"]: pe = 0
        if r["no_pb"]: pb = 0
        if r["no_ps"]: ps = 0
        if r["no_ev"]: ev = 0
        return pe, pb, ps, ev
    if code_upper in BANK_INSURANCE_CODES:
        ps, ev = 0, 0
    if code_upper in SPECIAL_EXCLUSIONS:
        for metric in SPECIAL_EXCLUSIONS[code_upper]:
            if metric == "PE": pe = 0
            elif metric == "PB": pb = 0
            elif metric == "PS": ps = 0
            elif metric == "EV": ev = 0
    return pe, pb, ps, ev

# ========== 公式 ==========

def row_formulas(n):
    return {
        "F": '=IF(OR(D{n}="",E{n}="",E{n}<=0,D{n}<=0),"",D{n}/E{n})'.format(n=n),
        "J": '=IF(OR(H{n}="",I{n}="",I{n}<=0,H{n}<=0),"",H{n}/I{n})'.format(n=n),
        "N": '=IF(OR(L{n}="",M{n}="",M{n}<=0,L{n}<=0),"",L{n}/M{n})'.format(n=n),
        "R": '=IF(OR(P{n}="",Q{n}="",Q{n}<=0,P{n}<=0),"",P{n}/Q{n})'.format(n=n),
        "U": '=IF(AND(ISNUMBER(D{n}),D{n}>0),1,0)'.format(n=n),
        "V": '=IF(AND(ISNUMBER(D{n}),ISNUMBER(H{n}),D{n}>0,H{n}>0),H{n}/D{n},"")'.format(n=n),
        "W": '=IF(AND(ISNUMBER(E{n}),ISNUMBER(I{n}),E{n}>0,I{n}>0),I{n}/E{n},"")'.format(n=n),
        "X": '=IF(AND(ISNUMBER(V{n}),ISNUMBER(W{n})),(V{n}-W{n})*100,"")'.format(n=n),
        "Y": '=IF(AND(G{n}=1,K{n}=1,ISNUMBER(V{n}),ISNUMBER(W{n})),1,0)'.format(n=n),
        "Z": '=IF(AND(ISNUMBER(H{n}),H{n}>0),1,0)'.format(n=n),
        "AA": '=IF(AND(G{n}=1,ISNUMBER(F{n}),F{n}>0),LN(F{n}),0)'.format(n=n),
        "AB": '=IF(AND(K{n}=1,ISNUMBER(J{n}),J{n}>0),LN(J{n}),0)'.format(n=n),
        "AC": '=IF(AND(O{n}=1,ISNUMBER(N{n}),N{n}>0),LN(N{n}),0)'.format(n=n),
        "AD": '=IF(AND(S{n}=1,ISNUMBER(R{n}),R{n}>0),LN(R{n}),0)'.format(n=n),
        "AE": '=IF(AND(G{n}=1,ISNUMBER(F{n}),F{n}>0),1/F{n},0)'.format(n=n),
        "AF": '=IF(AND(K{n}=1,ISNUMBER(J{n}),J{n}>0),1/J{n},0)'.format(n=n),
        "AG": '=IF(AND(O{n}=1,ISNUMBER(N{n}),N{n}>0),1/N{n},0)'.format(n=n),
        "AH": '=IF(AND(S{n}=1,ISNUMBER(R{n}),R{n}>0),1/R{n},0)'.format(n=n),
        "AI": '=IF(AND(G{n}=1,ISNUMBER(D{n}),D{n}>0),1/D{n},0)'.format(n=n),
        "AJ": '=IF(AND(G{n}=1,ISNUMBER(E{n}),E{n}>0),1/E{n},0)'.format(n=n),
        "AK": '=IF(AND(K{n}=1,ISNUMBER(H{n}),H{n}>0),1/H{n},0)'.format(n=n),
        "AL": '=IF(AND(K{n}=1,ISNUMBER(I{n}),I{n}>0),1/I{n},0)'.format(n=n),
        "AM": '=IF(AND(O{n}=1,ISNUMBER(L{n}),L{n}>0),1/L{n},0)'.format(n=n),
        "AN": '=IF(AND(O{n}=1,ISNUMBER(M{n}),M{n}>0),1/M{n},0)'.format(n=n),
        "AO": '=IF(AND(S{n}=1,ISNUMBER(P{n}),P{n}<>0,ISNUMBER(Q{n}),Q{n}>0),1/P{n},0)'.format(n=n),
        "AP": '=IF(AND(S{n}=1,ISNUMBER(P{n}),P{n}<>0,ISNUMBER(Q{n}),Q{n}>0),1/Q{n},0)'.format(n=n),
    }

# ========== 读取输入 ==========

def detect_columns(ws):
    headers = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(1, col).value
        if not val:
            continue
        text = str(val).strip().lower()
        for key, aliases in COL_ALIASES.items():
            if any(a.lower() in text or text in a.lower() for a in aliases):
                headers[key] = col
                break
    return headers

def read_holdings(ws):
    cols = detect_columns(ws)
    if "name" not in cols or "weight" not in cols:
        return []
    holdings = []
    for row in range(2, ws.max_row + 1):
        name = ws.cell(row, cols["name"]).value
        if not name:
            continue
        code = ws.cell(row, cols.get("code", 0)).value if "code" in cols else ""
        weight = ws.cell(row, cols["weight"]).value
        try:
            weight = float(weight) if weight is not None else 0.0
        except (TypeError, ValueError):
            weight = 0.0
        if weight > 1.5:
            weight = weight / 100.0
        holdings.append({
            "name": str(name).strip(),
            "code": str(code).strip() if code else "",
            "weight": weight,
        })
    holdings.sort(key=lambda x: x["weight"], reverse=True)
    return holdings

# ========== 生成计算页 ==========

def generate_sheet(ws, holdings, title="估值表", preferred_summary_start=None):
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "蓝字由 GF 更新器填入；进标志已按规则预填（可手工覆盖）。"
    ws["A2"].font = Font(italic=True, color="808080")

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    for col, h in enumerate(HEADERS, 1):
        cell = ws.cell(4, col, h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    start_row = 5
    for i, h in enumerate(holdings):
        r = start_row + i
        ws.cell(r, 1, h.get("name", ""))
        ws.cell(r, 2, h.get("code", ""))
        ws.cell(r, 3, h.get("weight", 0))
        pe, pb, ps, ev = get_inclusion_flags(h.get("code", ""))
        ws.cell(r, 7, pe)
        ws.cell(r, 11, pb)
        ws.cell(r, 15, ps)
        ws.cell(r, 19, ev)
        formulas = row_formulas(r)
        for col_letter, formula in formulas.items():
            ws[col_letter + str(r)] = formula

    last_data = start_row + len(holdings) - 1
    if last_data < start_row:
        last_data = start_row

    # 合计净值%：紧跟数据行之后
    total_row = last_data + 1
    ws.cell(total_row, 1, "合计净值%").font = Font(bold=True)
    ws.cell(total_row, 2, "=SUM(C{0}:C{1})".format(start_row, last_data))

    # 汇总区起点：优先用旧页的位置（更新时不挪动，总览引用才不会断）
    if preferred_summary_start and preferred_summary_start > total_row + 1:
        summary_start = preferred_summary_start
    else:
        summary_start = max(SUMMARY_START, total_row + 2)
    rng = "C{0}:C{1}".format(start_row, last_data)
    s, e = start_row, last_data

    summaries = [
        (summary_start, "PE 覆盖净值%", "=SUMPRODUCT({rng},G{s}:G{e})".format(rng=rng, s=s, e=e),
         "加权 PE/Med", '=IF(B{0}=0,"",SUMPRODUCT({rng},G{s}:G{e},F{s}:F{e})/B{0})'.format(summary_start, rng=rng, s=s, e=e)),
        (summary_start + 1, "PB 覆盖净值%", "=SUMPRODUCT({rng},K{s}:K{e})".format(rng=rng, s=s, e=e),
         "加权 PB/Med", '=IF(B{0}=0,"",SUMPRODUCT({rng},K{s}:K{e},J{s}:J{e})/B{0})'.format(summary_start + 1, rng=rng, s=s, e=e)),
        (summary_start + 2, "PS 覆盖净值%", "=SUMPRODUCT({rng},O{s}:O{e})".format(rng=rng, s=s, e=e),
         "加权 PS/Med", '=IF(B{0}=0,"",SUMPRODUCT({rng},O{s}:O{e},N{s}:N{e})/B{0})'.format(summary_start + 2, rng=rng, s=s, e=e)),
        (summary_start + 3, "EV/FCF 覆盖净值%", "=SUMPRODUCT({rng},S{s}:S{e})".format(rng=rng, s=s, e=e),
         "加权 EV/FCF/Med", '=IF(B{0}=0,"",SUMPRODUCT({rng},S{s}:S{e},R{s}:R{e})/B{0})'.format(summary_start + 3, rng=rng, s=s, e=e)),
        (summary_start + 5, "调和 PE（进调和PE=1）", None,
         None, '=IFERROR(SUMPRODUCT({rng},U{s}:U{e})/SUMPRODUCT({rng},U{s}:U{e},1/(D{s}:D{e}+(D{s}:D{e}<=0))),"")'.format(rng=rng, s=s, e=e)),
        (summary_start + 6, "持仓加权调和 EV/FCF", None,
         None, '=IFERROR(B{0}/SUMPRODUCT({rng},S{s}:S{e},1/(P{s}:P{e}+(P{s}:P{e}<=0))),"")'.format(summary_start + 3, rng=rng, s=s, e=e)),
        (summary_start + 8, "调和PE覆盖净值%（利润>0）", "=SUMPRODUCT({rng},U{s}:U{e})".format(rng=rng, s=s, e=e),
         "调和PE（同D）", "=D{0}".format(summary_start + 5)),
        (summary_start + 9, "调和PB覆盖净值%（账面>0）", "=SUMPRODUCT({rng},Z{s}:Z{e})".format(rng=rng, s=s, e=e),
         "调和PB", '=IFERROR(SUMPRODUCT({rng},Z{s}:Z{e})/SUMPRODUCT({rng},Z{s}:Z{e},1/(H{s}:H{e}+(H{s}:H{e}<=0))),"")'.format(rng=rng, s=s, e=e)),
        (summary_start + 10, "PE∩PB 交集覆盖%", "=SUMPRODUCT({rng},U{s}:U{e},Z{s}:Z{e})".format(rng=rng, s=s, e=e),
         "反推组合ROE=调和PB÷调和PE", '=IFERROR((SUMPRODUCT({rng},U{s}:U{e},Z{s}:Z{e})/SUMPRODUCT({rng},U{s}:U{e},Z{s}:Z{e},1/(H{s}:H{e}+(H{s}:H{e}<=0))))/(SUMPRODUCT({rng},U{s}:U{e},Z{s}:Z{e})/SUMPRODUCT({rng},U{s}:U{e},Z{s}:Z{e},1/(D{s}:D{e}+(D{s}:D{e}<=0)))),"")'.format(rng=rng, s=s, e=e)),
        (summary_start + 12, "几何 PE/Med", '=IF(B{0}=0,"",EXP(SUMPRODUCT({rng},AA{s}:AA{e})/B{0}))'.format(summary_start, rng=rng, s=s, e=e), None, None),
        (summary_start + 13, "几何 PB/Med", '=IF(B{0}=0,"",EXP(SUMPRODUCT({rng},AB{s}:AB{e})/B{0}))'.format(summary_start + 1, rng=rng, s=s, e=e), None, None),
        (summary_start + 14, "几何 PS/Med", '=IF(B{0}=0,"",EXP(SUMPRODUCT({rng},AC{s}:AC{e})/B{0}))'.format(summary_start + 2, rng=rng, s=s, e=e), None, None),
        (summary_start + 15, "几何 EV/FCF/Med", '=IF(B{0}=0,"",EXP(SUMPRODUCT({rng},AD{s}:AD{e})/B{0}))'.format(summary_start + 3, rng=rng, s=s, e=e), None, None),
        (summary_start + 17, "持仓加权 ΔROE（ppt）", '=IFERROR(SUMPRODUCT({rng},Y{s}:Y{e},X{s}:X{e})/SUMPRODUCT({rng},Y{s}:Y{e}),"")'.format(rng=rng, s=s, e=e),
         "反推ROE现−中位。解释列，不是贵贱。", None),
        (summary_start + 18, "净资产加权反推ROE", '=IFERROR(SUMPRODUCT({rng},U{s}:U{e},Z{s}:Z{e},IF(ISNUMBER(V{s}:V{e}),V{s}:V{e},0),1/H{s}:H{e})/SUMPRODUCT({rng},U{s}:U{e},Z{s}:Z{e},1/H{s}:H{e}),"")'.format(rng=rng, s=s, e=e),
         "应接近反推组合ROE。不要用持仓算术平均。", None),
        (summary_start + 20, "调和 PE/Med（市值回归）", '=IF(OR(B{0}=0,SUMPRODUCT({rng},AE{s}:AE{e})=0),"",B{0}/SUMPRODUCT({rng},AE{s}:AE{e}))'.format(summary_start, rng=rng, s=s, e=e), None, None),
        (summary_start + 21, "调和 PB/Med", '=IF(OR(B{0}=0,SUMPRODUCT({rng},AF{s}:AF{e})=0),"",B{0}/SUMPRODUCT({rng},AF{s}:AF{e}))'.format(summary_start + 1, rng=rng, s=s, e=e), None, None),
        (summary_start + 22, "调和 PS/Med", '=IF(OR(B{0}=0,SUMPRODUCT({rng},AG{s}:AG{e})=0),"",B{0}/SUMPRODUCT({rng},AG{s}:AG{e}))'.format(summary_start + 2, rng=rng, s=s, e=e), None, None),
        (summary_start + 23, "调和 EV/FCF/Med", '=IF(OR(B{0}=0,SUMPRODUCT({rng},AH{s}:AH{e})=0),"",B{0}/SUMPRODUCT({rng},AH{s}:AH{e}))'.format(summary_start + 3, rng=rng, s=s, e=e), None, None),
        (summary_start + 25, "当一家公司：组合PE/常态PE", '=IF(OR(B{0}=0,SUMPRODUCT({rng},AI{s}:AI{e})=0,SUMPRODUCT({rng},AJ{s}:AJ{e})=0),"",SUMPRODUCT({rng},AJ{s}:AJ{e})/SUMPRODUCT({rng},AI{s}:AI{e}))'.format(summary_start, rng=rng, s=s, e=e), None, None),
        (summary_start + 26, "当一家公司：组合PB/常态PB", '=IF(OR(B{0}=0,SUMPRODUCT({rng},AK{s}:AK{e})=0,SUMPRODUCT({rng},AL{s}:AL{e})=0),"",SUMPRODUCT({rng},AL{s}:AL{e})/SUMPRODUCT({rng},AK{s}:AK{e}))'.format(summary_start + 1, rng=rng, s=s, e=e), None, None),
        (summary_start + 27, "当一家公司：组合PS/常态PS", '=IF(OR(B{0}=0,SUMPRODUCT({rng},AM{s}:AM{e})=0,SUMPRODUCT({rng},AN{s}:AN{e})=0),"",SUMPRODUCT({rng},AN{s}:AN{e})/SUMPRODUCT({rng},AM{s}:AM{e}))'.format(summary_start + 2, rng=rng, s=s, e=e), None, None),
        (summary_start + 28, "当一家公司：组合EVFCF/常态EVFCF", '=IF(OR(SUMPRODUCT({rng},AO{s}:AO{e})<=0,SUMPRODUCT({rng},AP{s}:AP{e})=0),"整包FCF≤0",SUMPRODUCT({rng},AP{s}:AP{e})/SUMPRODUCT({rng},AO{s}:AO{e}))'.format(rng=rng, s=s, e=e), None, None),
        (summary_start + 30, "进PE/Med样本上的调和PE", '=IF(OR(B{0}=0,SUMPRODUCT({rng},AI{s}:AI{e})=0),"",B{0}/SUMPRODUCT({rng},AI{s}:AI{e}))'.format(summary_start, rng=rng, s=s, e=e), None, None),
        (summary_start + 31, "进PE/Med样本上的调和中位PE（常态）", '=IF(OR(B{0}=0,SUMPRODUCT({rng},AJ{s}:AJ{e})=0),"",B{0}/SUMPRODUCT({rng},AJ{s}:AJ{e}))'.format(summary_start, rng=rng, s=s, e=e), None, None),
    ]

    notes = [
        (summary_start + 33, "口径说明（客观）"),
        (summary_start + 34, "加权：Σ(w×r)/Σw，样本内算术平均。"),
        (summary_start + 35, "几何：exp(Σ(w×ln r)/Σw)，对极端值更稳健。"),
        (summary_start + 36, "调和：Σw/Σ(w/r)，等价于利润不变、各回中位时的市值回归比。"),
        (summary_start + 37, "当一家公司：Σ(w/中位)/Σ(w/现)，把组合当成一家公司看相对自身常态。"),
        (summary_start + 38, "银行/保险/券商默认不进 PS、EV（见 gf_ticker_map）。"),
        (summary_start + 39, "负 PE / 负账面 / 无中位 的项自动不进对应指标。"),
    ]

    for row, a, b, c, d in summaries:
        if a:
            ws.cell(row, 1, a).font = Font(bold=True)
        if b:
            ws.cell(row, 2, b)
        if c:
            ws.cell(row, 3, c)
        if d:
            ws.cell(row, 4, d)

    for row, text in notes:
        ws.cell(row, 1, text).font = Font(italic=True, color="666666")

    # 历史流水区（只建表头，数据行由更新器 snapshot_as_one_company 写入）
    hist_start = 105
    ws.cell(hist_start, 1, "当一家公司历史").font = Font(bold=True)
    ws.cell(hist_start + 1, 1, "时间")
    ws.cell(hist_start + 1, 2, "组合PE/常态PE")
    ws.cell(hist_start + 1, 3, "组合PB/常态PB")
    ws.cell(hist_start + 1, 4, "组合PS/常态PS")
    ws.cell(hist_start + 1, 5, "组合EVFCF/常态EVFCF")
    for col in range(1, 6):
        ws.cell(hist_start + 1, col).font = Font(bold=True, color="FFFFFF")
        ws.cell(hist_start + 1, col).fill = PatternFill("solid", fgColor="70AD47")

    # 散点图：当一家公司历史四维（初始范围，更新器会扩展）
    from openpyxl.chart import ScatterChart, Reference, Series
    from openpyxl.chart.series import SeriesLabel
    chart = ScatterChart()
    chart.title = "当一家公司历史"
    chart.style = 10
    chart.x_axis.title = "时间"
    chart.y_axis.title = "相对中位"
    chart.x_axis.numFmt = "MM-DD"
    chart.width = 18
    chart.height = 8
    xvalues = Reference(ws, min_col=1, min_row=107, max_row=108)
    titles = ["组合PE/常态PE", "组合PB/常态PB", "组合PS/常态PS", "组合EVFCF/常态EVFCF"]
    for col, title in enumerate(titles, 2):
        yvalues = Reference(ws, min_col=col, min_row=106, max_row=108)
        s = Series(yvalues, xvalues, title_from_data=True)
        s.title = SeriesLabel(v=title)
        chart.series.append(s)
    ws.add_chart(chart, "A110")

    widths = {"A": 16, "B": 10, "C": 12, "T": 30}
    for col in range(1, len(HEADERS) + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = widths.get(letter, 12)
    ws.freeze_panes = "A5"
    return {"last_data_row": last_data, "summary_start": summary_start, "holdings_count": len(holdings)}

# ========== 生成总览 ==========

def generate_overview(wb, sheet_infos):
    ws = wb.create_sheet("总览对比", 0)
    ws["A1"] = "权益 QDII / ETF 估值快照（自动生成）"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "指标行号已对齐主文件「总览对比」（第13行起）。可直接把新列复制进主文件对应行。"
    ws["A2"].font = Font(italic=True, color="808080")

    # 表头
    ws.cell(4, 1, "项目")
    # 描述行占位（主文件有基金性质/持仓时点等，这里留空方便对齐）
    for r in [5, 6, 7, 8, 9]:
        ws.cell(r, 1, "")
    ws.cell(11, 1, "前十可覆盖样本（公式链到分表）").font = Font(bold=True)
    ws.cell(12, 1, "指标").font = Font(bold=True)

    for i, (name, _) in enumerate(sheet_infos):
        ws.cell(4, i + 2, name)

    # 完整指标行（行号对齐主文件「总览对比」，方便把整列复制进主文件）
    # 公式指向各分表自己的汇总区（summary_start 起）
    # 主文件总览行号：13=PE覆盖, 14=加权PE, 15=调和PE, 16=PB覆盖, ...
    metrics = [
        (13, "PE 覆盖净值%", "B", 0),
        (14, "加权 PE / 十年中位", "D", 0),
        (15, "调和 PE（利润>0）", "D", 5),
        (16, "PB 覆盖净值%", "B", 1),
        (17, "加权 PB / 十年中位", "D", 1),
        (18, "PS 覆盖净值%", "B", 2),
        (19, "加权 PS / 十年中位", "D", 2),
        (20, "EV/FCF 覆盖净值%", "B", 3),
        (21, "加权 EV/FCF / 十年中位", "D", 3),
        (22, "调和 EV/FCF（正FCF）", "D", 6),
        (23, "调和PE覆盖%（利润>0）", "B", 8),
        (24, "调和 PB", "D", 9),
        (25, "调和PB覆盖%（账面>0）", "B", 9),
        (26, "反推组合ROE（交集）", "D", 10),
        (27, "几何 PE/Med", "B", 12),
        (28, "几何 PB/Med", "B", 13),
        (29, "持仓加权 ΔROE（ppt）", "B", 17),
        (30, "几何 PS/Med", "B", 14),
        (31, "几何 EV/FCF/Med", "B", 15),
        (32, "调和 PE/Med（市值回归）", "B", 20),
        (33, "调和 PB/Med", "B", 21),
        (34, "调和 PS/Med", "B", 22),
        (35, "调和 EV/FCF/Med", "B", 23),
        (45, "当一家公司：组合PE/常态PE", "B", 25),
        (46, "当一家公司：组合PB/常态PB", "B", 26),
        (47, "当一家公司：组合PS/常态PS", "B", 27),
        (48, "当一家公司：组合EVFCF/常态EVFCF", "B", 28),
    ]

    for row, label, col_letter, offset in metrics:
        ws.cell(row, 1, label).font = Font(bold=True)
        for i, (name, info) in enumerate(sheet_infos):
            safe = "'{}'".format(name) if any(c in name for c in " -") else name
            ss = info.get("summary_start", SUMMARY_START)
            ref_row = ss + offset
            ws.cell(row, i + 2, "={}!{}{}".format(safe, col_letter, ref_row))

    # 图表数据区（从 R54 起，对齐 update_gf_valuations.py 的 refresh_overview_rank）
    chart_start = 54
    ws.cell(chart_start - 2, 1, "估值相对十年中位对比（当一家公司四维；<1 便宜）").font = Font(bold=True)
    headers = ["基金", "PE/常态", "PB/常态", "PS/常态", "EV/常态", "便宜区(<1)", "贵区(>1)"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(chart_start - 1, col, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")

    for i, (name, _) in enumerate(sheet_infos):
        row = chart_start + i
        safe = "'{}'".format(name) if any(c in name for c in " -") else name
        ws.cell(row, 1, name)
        col_letter = get_column_letter(i + 2)
        ws.cell(row, 2, "={}".format(col_letter + "45"))
        ws.cell(row, 3, "={}".format(col_letter + "46"))
        ws.cell(row, 4, "={}".format(col_letter + "47"))
        ws.cell(row, 5, "={}".format(col_letter + "48"))
        ws.cell(row, 6, 1)
        ws.cell(row, 7, 1.5)

    # 不在这里画图；跑 update_gf_valuations.py 时会调用 refresh_overview_rank 重建图表

    # 列宽
    ws.column_dimensions["A"].width = 26
    for i in range(len(sheet_infos)):
        ws.column_dimensions[get_column_letter(i + 2)].width = 18
    ws.freeze_panes = "B5"

    # 列宽
    ws.column_dimensions["A"].width = 26
    for i in range(len(sheet_infos)):
        ws.column_dimensions[get_column_letter(i + 2)].width = 18
    ws.freeze_panes = "B5"

# ========== 生成辅助页 ==========

def generate_sources_sheet(wb):
    ws = wb.create_sheet("网页更新来源")
    ws["A1"] = "网页更新来源记录"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "下次更新持仓时优先使用以下来源。"
    ws["A2"].font = Font(italic=True, color="808080")
    ws["A4"] = "代码/简称"
    ws["B4"] = "来源说明"
    ws["C4"] = "网址"
    for col in range(1, 4):
        ws.cell(4, col).font = Font(bold=True, color="FFFFFF")
        ws.cell(4, col).fill = PatternFill("solid", fgColor="4472C4")
    for i, (code, desc, url) in enumerate(SOURCES):
        ws.cell(5 + i, 1, code)
        ws.cell(5 + i, 2, desc)
        ws.cell(5 + i, 3, url)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 70

def generate_methodology_sheet(wb):
    ws = wb.create_sheet("口径说明")
    ws["A1"] = "口径说明（客观）"
    ws["A1"].font = Font(bold=True, size=14)
    notes = [
        (3, "相对中位", "各指标 = 现值 / 十年中位。中位来自 GuruFocus。"),
        (4, "加权", "Σ(w × r) / Σw。样本内算术平均，对极端值敏感。"),
        (5, "几何", "exp(Σ(w × ln r) / Σw)。对极端值更稳健。"),
        (6, "调和", "Σw / Σ(w / r)。等价于利润不变、各回中位时的市值回归比。"),
        (7, "当一家公司", "Σ(w / 中位) / Σ(w / 现)。把组合当成一家公司看相对自身常态。"),
        (8, "进标志", "银行/保险/券商默认不进 PS、EV（见 gf_ticker_map.csv 的 no_ps/no_ev）。"),
        (9, "负值处理", "负 PE / 负账面 / 无中位 的项自动不进对应指标。"),
        (10, "EV/FCF", "负 FCF 时 1/现为负，整包收益率被拉低；中位仍须 >0。"),
        (11, "覆盖率", "进入该指标的持仓权重之和。未进入的不参与加权/几何/调和。"),
        (12, "历史流水", "每次生成追加一行，公式链到汇总区。蓝字更新后历史自动刷新。"),
    ]
    for row, title, text in notes:
        ws.cell(row, 1, title).font = Font(bold=True)
        ws.cell(row, 2, text)
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 70

def generate_master_history(wb, sheet_infos):
    ws = wb.create_sheet("当一家公司历史")
    ws["A1"] = "当一家公司历史总账"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "各页历史区最新快照的引用。蓝字更新后自动刷新。"
    ws["A2"].font = Font(italic=True, color="808080")
    headers = ["页名", "最新时间", "组合PE/常态PE", "组合PB/常态PB", "组合PS/常态PS", "组合EVFCF/常态EVFCF"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(4, col, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="70AD47")
    for i, (name, _) in enumerate(sheet_infos):
        row = 5 + i
        safe = "'{}'".format(name) if any(c in name for c in " -") else name
        ws.cell(row, 1, name)
        ws.cell(row, 2, "={0}!A107".format(safe))
        ws.cell(row, 3, "={0}!B107".format(safe))
        ws.cell(row, 4, "={0}!C107".format(safe))
        ws.cell(row, 5, "={0}!D107".format(safe))
        ws.cell(row, 6, "={0}!E107".format(safe))
    for col in range(1, 7):
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws.column_dimensions["A"].width = 22

# ========== 主流程 ==========

def generate_from_input(input_path, output_path):
    if not os.path.exists(input_path):
        print("输入文件不存在:", input_path)
        return False

    src = openpyxl.load_workbook(input_path, data_only=True)
    out_wb = openpyxl.Workbook()
    default = out_wb.active
    out_wb.remove(default)

    sheet_infos = []
    for sheet_name in src.sheetnames:
        src_ws = src[sheet_name]
        holdings = read_holdings(src_ws)
        if not holdings:
            print("跳过空 sheet:", sheet_name)
            continue
        out_name = sheet_name[:31]
        out_ws = out_wb.create_sheet(out_name)
        info = generate_sheet(out_ws, holdings, title=sheet_name)
        sheet_infos.append((out_name, info))
        print("已生成:", out_name, "持仓", info["holdings_count"], "只")

    if not sheet_infos:
        print("没有生成任何 sheet。")
        return False

    generate_overview(out_wb, sheet_infos)
    generate_sources_sheet(out_wb)
    generate_methodology_sheet(out_wb)
    generate_master_history(out_wb, sheet_infos)
    print("已生成: 总览对比 + 网页更新来源 + 口径说明 + 当一家公司历史")

    out_wb.save(output_path)
    print("\n完成。输出文件:", output_path)
    print("共生成", len(sheet_infos), "个计算页 + 4 个辅助页。")
    return True

def create_sample_input(path):
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "QQQ_示例"
    ws1["A1"] = "名称"
    ws1["B1"] = "代码"
    ws1["C1"] = "净值权重"
    sample = [
        ("英伟达", "NVDA", 0.0820),
        ("苹果", "AAPL", 0.0791),
        ("高盛", "GS", 0.0500),
        ("联合健康", "UNH", 0.0400),
    ]
    for i, (n, c, w) in enumerate(sample, 2):
        ws1.cell(i, 1, n)
        ws1.cell(i, 2, c)
        ws1.cell(i, 3, w)
    wb.save(path)
    print("已生成示例输入:", path)

def ensure_xlsx(name):
    name = name.strip().strip('"').strip("'")
    if not name:
        return ""
    return name if name.lower().endswith(".xlsx") else name + ".xlsx"


def _read_blue_values(ws):
    """从旧计算页读取蓝字（现价/中位），按代码索引，更新后写回。"""
    blues = {}
    # 蓝字列：D E H I L M P Q
    blue_cols = [4, 5, 8, 9, 12, 13, 16, 17]
    for r in range(5, ws.max_row + 1):
        code = ws.cell(r, 2).value
        if not code:
            continue
        code = str(code).strip().upper()
        vals = {}
        for c in blue_cols:
            v = ws.cell(r, c).value
            if v is not None and not (isinstance(v, str) and v.startswith("=")):
                vals[c] = v
        if vals:
            blues[code] = vals
    return blues


def _find_summary_start(ws):
    """找旧页「PE 覆盖净值%」所在行，作为汇总区起点。"""
    for r in range(20, min(ws.max_row + 1, 150)):
        a = ws.cell(r, 1).value
        if a and str(a).startswith("PE 覆盖"):
            return r
    return None


def _copy_sheet_content(src_ws, dest):
    for row in src_ws.iter_rows():
        for cell in row:
            dest.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                dest.cell(cell.row, cell.column).font = cell.font.copy()
                dest.cell(cell.row, cell.column).fill = cell.fill.copy()
                dest.cell(cell.row, cell.column).alignment = cell.alignment.copy()
                dest.cell(cell.row, cell.column).number_format = cell.number_format
    for col_letter, dim in src_ws.column_dimensions.items():
        dest.column_dimensions[col_letter].width = dim.width
    for ch in src_ws._charts:
        dest.add_chart(ch)


def _restore_blues(ws, blues):
    """按代码把旧蓝字写回新页。"""
    if not blues:
        return 0
    restored = 0
    for r in range(5, ws.max_row + 1):
        code = ws.cell(r, 2).value
        if not code:
            continue
        key = str(code).strip().upper()
        if key in blues:
            for c, v in blues[key].items():
                ws.cell(r, c, v)
            restored += 1
    return restored


def merge_into_main(input_path, main_path, sheet_names=None):
    """把输入里的持仓合并进主文件。
    - 更新（主文件已有同名计算页）：只重生该计算页（公式范围对齐、补合计净值%、保留蓝字和汇总区位置）。总览对比、排名表、历史总账一律不动。
    - 添加（主文件没有该页）：新建计算页，并在总览追加列、排名表插入行、历史总账追加。
    """
    if not os.path.exists(main_path):
        print("主文件不存在:", main_path)
        return False

    src = openpyxl.load_workbook(input_path, data_only=True)
    main = openpyxl.load_workbook(main_path)
    overview = main["总览对比"] if "总览对比" in main.sheetnames else None

    updated = 0
    added = 0

    for sheet_name in src.sheetnames:
        if sheet_names and sheet_name not in sheet_names:
            continue
        src_ws = src[sheet_name]
        holdings = read_holdings(src_ws)
        if not holdings:
            continue
        out_name = sheet_name[:31]

        if out_name in main.sheetnames:
            # ========== 更新模式：只动计算页 ==========
            old_ws = main[out_name]
            blues = _read_blue_values(old_ws)
            old_summary = _find_summary_start(old_ws)

            # 生成到临时簿
            tmp = openpyxl.Workbook()
            tmp.remove(tmp.active)
            tmp_ws = tmp.create_sheet(out_name)
            info = generate_sheet(tmp_ws, holdings, title=sheet_name,
                                  preferred_summary_start=old_summary)
            restored = _restore_blues(tmp_ws, blues)

            # 替换主文件里的页
            del main[out_name]
            dest = main.create_sheet(out_name)
            _copy_sheet_content(tmp_ws, dest)

            note = ""
            if old_summary and info["summary_start"] != old_summary:
                note = "（汇总区从第{}行移到第{}行，总览引用可能需手工对齐）".format(
                    old_summary, info["summary_start"])
            print("已更新计算页: {} 持仓{}只 蓝字回写{}只{}".format(
                out_name, info["holdings_count"], restored, note))
            print("  → 总览对比 / 排名表 / 历史总账未改动")
            updated += 1
        else:
            # ========== 添加模式：新建页 + 动总览 ==========
            if overview is None:
                print("主文件没有「总览对比」，无法添加新页:", out_name)
                continue
            tmp = openpyxl.Workbook()
            tmp.remove(tmp.active)
            tmp_ws = tmp.create_sheet(out_name)
            info = generate_sheet(tmp_ws, holdings, title=sheet_name)

            dest = main.create_sheet(out_name)
            _copy_sheet_content(tmp_ws, dest)

            # 总览追加列
            last_col = 1
            for c in range(1, overview.max_column + 2):
                if overview.cell(4, c).value:
                    last_col = c
            next_col = last_col + 1
            overview.cell(4, next_col, out_name)
            ss = info.get("summary_start", SUMMARY_START)
            safe = "'{}'".format(out_name) if any(c in out_name for c in " -") else out_name
            metrics = [
                (13, "B", 0), (14, "D", 0), (15, "D", 5),
                (16, "B", 1), (17, "D", 1), (18, "B", 2), (19, "D", 2),
                (20, "B", 3), (21, "D", 3), (22, "D", 6),
                (23, "B", 8), (24, "D", 9), (25, "B", 9), (26, "D", 10),
                (27, "B", 12), (28, "B", 13), (29, "B", 17),
                (30, "B", 14), (31, "B", 15),
                (32, "B", 20), (33, "B", 21), (34, "B", 22), (35, "B", 23),
                (45, "B", 25), (46, "B", 26), (47, "B", 27), (48, "B", 28),
            ]
            for row, col_letter, offset in metrics:
                overview.cell(row, next_col, "={}!{}{}".format(safe, col_letter, ss + offset))

            # 排名表插入
            note_row = None
            for r in range(54, 120):
                v = overview.cell(r, 1).value
                if v and "横坐标已按四维平均" in str(v):
                    note_row = r
                    break
            if note_row is None:
                rank_row = 54
                while overview.cell(rank_row, 1).value:
                    rank_row += 1
            else:
                rank_row = note_row
                overview.insert_rows(note_row)
            col_letter = get_column_letter(next_col)
            overview.cell(rank_row, 1, out_name)
            overview.cell(rank_row, 2, "={}45".format(col_letter))
            overview.cell(rank_row, 3, "={}46".format(col_letter))
            overview.cell(rank_row, 4, "={}47".format(col_letter))
            overview.cell(rank_row, 5, "={}48".format(col_letter))
            overview.cell(rank_row, 6, 1)
            overview.cell(rank_row, 7, 1.5)
            overview.cell(rank_row, 8, '=IF(COUNT(B{0}:E{0})=0,"",AVERAGE(B{0}:E{0}))'.format(rank_row))

            # 历史总账追加
            if "当一家公司历史" in main.sheetnames:
                hist = main["当一家公司历史"]
                hist_row = 5
                while hist.cell(hist_row, 1).value:
                    hist_row += 1
                hist.cell(hist_row, 1, out_name)
                hist.cell(hist_row, 2, "={}!A107".format(safe))
                hist.cell(hist_row, 3, "={}!B107".format(safe))
                hist.cell(hist_row, 4, "={}!C107".format(safe))
                hist.cell(hist_row, 5, "={}!D107".format(safe))
                hist.cell(hist_row, 6, "={}!E107".format(safe))

            print("已添加新页: {} 持仓{}只，并在总览第{}列追加".format(
                out_name, info["holdings_count"], next_col))
            added += 1

    # 网页更新来源追加（仅添加模式可能需要，更新不动）
    if added and "网页更新来源" in main.sheetnames:
        src_sheet = main["网页更新来源"]
        existing = set()
        for r in range(5, src_sheet.max_row + 1):
            v = src_sheet.cell(r, 1).value
            if v:
                existing.add(str(v))
        for code, desc, url in SOURCES:
            if code not in existing:
                row = src_sheet.max_row + 1
                src_sheet.cell(row, 1, code)
                src_sheet.cell(row, 2, desc)
                src_sheet.cell(row, 3, url)
                print("已追加来源:", code)

    main.save(main_path)
    print("\n完成。更新 {} 页，添加 {} 页。主文件: {}".format(updated, added, main_path))
    return True



def interactive_run():
    print("=" * 60)
    print("估值表生成器（交互模式）")
    print("=" * 60)
    print()
    print("输入文件要求：")
    print("  • Excel 文件，每个 sheet 代表一个指数/组合")
    print("  • 每个 sheet 只需三列：名称 | 代码 | 净值权重")
    print("  • 权重可以是小数（0.082）或百分数（8.2）")
    print("  • 第一行为表头，数据从第二行开始")
    print()
    print("如果你还没有输入文件，可以先生成模板。")
    print()

    while True:
        raw = input("请输入持仓文件路径（直接回车则生成模板 持仓输入.xlsx）: ").strip()
        if not raw:
            print()
            print("正在生成模板 持仓输入.xlsx ...")
            create_sample_input("持仓输入.xlsx")
            print()
            print("模板已生成。请用 Excel 打开 持仓输入.xlsx，")
            print("把你的持仓填进任意 sheet（保留三列表头），保存后重新运行本脚本。")
            print()
            again = input("是否现在就用刚生成的模板继续？(y/N): ").strip().lower()
            if again in ("y", "yes", "是"):
                input_path = "持仓输入.xlsx"
                break
            else:
                return
        else:
            input_path = ensure_xlsx(raw)
            if os.path.exists(input_path):
                break
            print("文件不存在:", input_path)
            gen = input("是否生成模板到该路径？(y/N): ").strip().lower()
            if gen in ("y", "yes", "是"):
                create_sample_input(input_path)
                print("模板已生成。请填好后重新运行。")
                return
            print("请重新输入正确路径。\n")

    default_out = "估值表输出.xlsx"
    raw_out = input("输出文件路径（直接回车使用默认 {}）: ".format(default_out)).strip()
    output_path = ensure_xlsx(raw_out) if raw_out else default_out

    print()
    print("输入:", input_path)
    print("输出:", output_path)
    print()
    generate_from_input(input_path, output_path)

    print()
    print("合并规则：主文件已有同名计算页 → 只更新该页（公式对齐、补合计、保留蓝字），总览不动；")
    print("          主文件没有该页 → 新建计算页，并在总览追加列。")
    merge = input("是否直接合并进主文件「权益QDII_ETF_估值相对十年中位.xlsx」？(y/N): ").strip().lower()
    if merge in ("y", "yes", "是"):
        main_path = input("主文件路径（直接回车用默认）: ").strip()
        if not main_path:
            main_path = "权益QDII_ETF_估值相对十年中位.xlsx"
        main_path = ensure_xlsx(main_path)
        merge_into_main(input_path, main_path)

if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "--sample":
        create_sample_input("持仓输入.xlsx")
    elif len(sys.argv) >= 2:
        # 命令行模式：python generate_valuation_sheet.py 输入.xlsx [输出.xlsx]
        input_path = ensure_xlsx(sys.argv[1])
        output_path = ensure_xlsx(sys.argv[2]) if len(sys.argv) >= 3 else "估值表输出.xlsx"
        generate_from_input(input_path, output_path)
    else:
        # 无参数：进入交互模式
        interactive_run()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从 GuruFocus 公开 term 页抓 PE/PB/PS/EV-FCF 的 Current + 十年 Med，
写回《权益QDII_ETF_估值相对十年中位.xlsx》蓝字格。

不改权重、不改加权/调和/当一家公司公式。
需要本机 Chrome（由 playwright install chromium 提供）。

对照表 gf_ticker_map.csv 优先于后缀猜测。
2026-09-23：印尼兜底 ISX（不再只用 IDX），马来 XKLS，巴西无后缀时试 BSP:。
沙特四位代码必须写在对照表（SAU:1120），脚本不自动猜。
2026-10-09 修复版：
  - 被拦截（Sorry you have been blocked / Attention Required）后不写备注、不追加快照、不保存 xlsx，后面的页不再跑。
  - 某行四项都没取到则整行不动；本页没写回任何一行就不追加快照；零写回不保存。
  - 不再每批 taskkill 调试 Chrome、不再隐身、不再换空配置；在已开的 9222 窗口里开标签，读完只关本批标签。
  - Just a moment / 正在进行安全验证 不算拦截，首次出现后至少再等 30 秒。
  - 修复缓存为空时永远存不进去的问题（旧版日志一直「缓存 0 条」）。
2026-10-15：
  - no_ps / no_ev（银行、券商、保险等）不再生成待开 URL，也不提示开页。
  - PE/PB/PS 现价或中位 ≤ 0 时，「进」仍由 include_flag 自动写 0。
  - EV/FCF：负数也进（加权、调和、当一家公司都计入）；进不进只看 no_ev（银行/券商）。
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import random
import re
import subprocess
import sys
import time
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.chart import AreaChart, BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.data_source import AxDataSource, NumData, NumRef, NumVal, StrRef
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Font
from openpyxl.utils import column_index_from_string

HERE = Path(__file__).resolve().parent
XLSX_NAME = "权益QDII_ETF_估值相对十年中位.xlsx"
# 优先用和脚本放在同一文件夹的表；没有再找上一级
DEFAULT_XLSX = HERE / XLSX_NAME if (HERE / XLSX_NAME).exists() else HERE.parent / XLSX_NAME
MAP_CSV = HERE / "gf_ticker_map.csv"
CACHE_JSON = HERE / "gf_cache.json"
CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium",
]
# 脚本自己开调试 Chrome 时用的固定配置目录：Cloudflare 放行 cookie 留在这里，下次还能用。
# 不再每批换空目录、不再用隐身。
CDP_PROFILE = HERE / "gf_cdp_profile"
ACTIVE_PROFILE: Path | None = None


def find_chrome() -> str:
    for p in CHROME_CANDIDATES:
        if Path(p).exists():
            return p
    return CHROME_CANDIDATES[0]


def find_edge() -> str:
    for p in (
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ):
        if Path(p).exists():
            return p
    return ""


def chrome_cmd(extra_urls: list[str] | None = None) -> list[str]:
    global ACTIVE_PROFILE
    if ACTIVE_PROFILE is None:
        ACTIVE_PROFILE = CDP_PROFILE
    ACTIVE_PROFILE.mkdir(parents=True, exist_ok=True)
    cmd = [
        find_chrome(),
        "--remote-debugging-port=9222",
        "--remote-allow-origins=*",
        f"--user-data-dir={str(ACTIVE_PROFILE)}",
    ]
    if extra_urls:
        cmd.extend(extra_urls)
    return cmd


DEBUG_CHROME_PROC = None


def _port_open() -> bool:
    try:
        import urllib.request

        urllib.request.urlopen("http://127.0.0.1:9222/json/version", timeout=1)
        return True
    except Exception:
        return False


def kill_debug_chrome() -> None:
    """只关带 9222 的调试窗，不动日常 Chrome。"""
    global DEBUG_CHROME_PROC
    if os.name == "nt":
        subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "Get-CimInstance Win32_Process | "
                "Where-Object { $_.Name -match 'chrome.exe|msedge.exe' -and $_.CommandLine -like '*remote-debugging-port=9222*' } | "
                "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }",
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    else:
        subprocess.run(["pkill", "-f", "remote-debugging-port=9222"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if DEBUG_CHROME_PROC is not None:
        try:
            DEBUG_CHROME_PROC.kill()
        except Exception:
            pass
        DEBUG_CHROME_PROC = None
    for _ in range(20):
        if not _port_open():
            return
        time.sleep(0.25)


def edge_user_data() -> Path:
    local = os.environ.get("LOCALAPPDATA") or ""
    return Path(local) / "Microsoft" / "Edge" / "User Data"


def launch_in_daily_edge(url: str) -> None:
    """只开一次脚本自己的窗口，用来放待开.html。不再开隐身，回车时也不再调用。"""
    global DEBUG_CHROME_PROC
    if _port_open():
        return
    browser = find_edge() or find_chrome()
    profile = HERE / "gf_click_profile"
    profile.mkdir(parents=True, exist_ok=True)
    print(f"只开这一只窗口。请在这只里面 Ctrl+单击，不要用日常 Edge。")
    DEBUG_CHROME_PROC = subprocess.Popen(
        [
            browser,
            "--remote-debugging-port=9222",
            "--remote-allow-origins=*",
            f"--user-data-dir={str(profile)}",
            url,
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    for _ in range(25):
        time.sleep(0.4)
        if _port_open():
            print("这只窗口已可读取。回车不会再开新浏览器。")
            return
    print("窗口没起来。")


def _cdp_json(path: str, method: str = "GET"):
    import urllib.request

    req = urllib.request.Request(f"http://127.0.0.1:9222{path}", method=method)
    with urllib.request.urlopen(req, timeout=8) as resp:
        raw = resp.read().decode("utf-8", "ignore")
    if not raw:
        return None
    try:
        return json.loads(raw)
    except Exception:
        return raw


def cdp_open_tab(url: str) -> bool:
    """在已经挂着 9222 的那只 Chrome 里开标签，不再另起隐身窗口。"""
    from urllib.parse import quote

    enc = quote(url, safe="")
    for method in ("PUT", "GET"):
        try:
            _cdp_json(f"/json/new?{enc}", method=method)
            return True
        except Exception:
            continue
    return False


def cdp_list_tabs() -> list[dict]:
    try:
        data = _cdp_json("/json/list")
    except Exception:
        return []
    return data if isinstance(data, list) else []


def cdp_close_tabs(urls: list[str]) -> int:
    want = {u.rstrip("/").lower() for u in urls}
    closed = 0
    for tab in cdp_list_tabs():
        u = (tab.get("url") or "").rstrip("/").lower()
        if u not in want:
            continue
        tid = tab.get("id")
        if not tid:
            continue
        try:
            _cdp_json(f"/json/close/{tid}")
            closed += 1
        except Exception:
            continue
    return closed


def open_one_in_connected_window(client, url: str) -> bool:
    """只在已经连上的 9222 窗口里开一个标签。不再另起 chrome.exe。"""
    print(f"      {url}")
    page = None
    for p in client.iter_pages():
        page = p
        break
    if page is not None:
        try:
            page.evaluate("(u) => window.open(u, '_blank')", url)
            time.sleep(1.5)
            tabs = cdp_list_tabs()
            if any(url.split("/")[-1].lower() in (t.get("url") or "").lower() for t in tabs):
                print("    已在 9222 这只窗口新开标签")
                return True
        except Exception as e:
            print(f"    从当前页打开失败，改用调试口：{e}")
    return cdp_open_tab(url)

SKIP_SHEETS = {"总览对比", "银行股息对照", "口径说明", "当一家公司历史", "网页更新来源"}
HIST_SHEET = "当一家公司历史"
HIST_COLS = [
    "时间",
    "页名",
    "组合PE/常态PE",
    "组合PB/常态PB",
    "组合PS/常态PS",
    "组合EVFCF/常态EVFCF",
    "来源",
]

def norm_ticker(t: str) -> str:
    s = (t or "").strip()
    try:
        from urllib.parse import unquote

        s = unquote(s)
    except Exception:
        pass
    s = s.replace("%3A", ":").replace("%3a", ":")
    if "/" in s:
        s = s.rstrip("/").split("/")[-1]
    return s


def ticker_aliases(t: str) -> list[str]:
    raw = norm_ticker(t)
    out = []
    for v in (raw, raw.upper(), raw.lower()):
        if v and v not in out:
            out.append(v)
    if ":" in raw:
        tail = raw.split(":")[-1]
        for v in (tail, tail.upper(), tail.lower()):
            if v and v not in out:
                out.append(v)
    return out


def cache_lookup(cache: dict | None, ticker: str, slug: str):
    if not cache:
        return None
    for t in ticker_aliases(ticker):
        hit = cache.get(f"{t}|{slug}")
        if hit:
            return hit
    return None


def cache_store(cache: dict, ticker: str, slug: str, parsed: dict) -> None:
    if cache is None or not parsed:
        return
    parsed = dict(parsed)
    parsed["fetched_at"] = date.today().isoformat()
    for t in ticker_aliases(ticker):
        cache[f"{t}|{slug}"] = parsed


def cache_usable(hit: dict | None, max_days: int = 15) -> bool:
    """半个月内取过的都算，现价是 None 也算，不再反复查。"""
    if not hit:
        return False
    ts = hit.get("fetched_at")
    if ts:
        try:
            age = (date.today() - date.fromisoformat(str(ts)[:10])).days
            if age > max_days:
                return False
        except Exception:
            pass
    return "med" in hit or "current" in hit


TERM_SPECS = [
    # field, column_current, column_median, column_include, url slugs to try
    ("pe", 4, 5, 7, ["pettm"]),
    ("pb", 8, 9, 11, ["pb-ratio", "pb"]),
    ("ps", 12, 13, 15, ["ps-ratio", "ps"]),
    ("ev", 16, 17, 19, ["enterprise-value-to-fcf", "ev2fcf"]),
]

RANGE_RE = re.compile(
    r"Min:\s*([+-]?(?:\d+(?:\.\d+)?|N/?A|At Loss))\s*"
    r"Med:\s*([+-]?(?:\d+(?:\.\d+)?|N/?A|At Loss))\s*"
    r"Max:\s*([+-]?(?:\d+(?:\.\d+)?|N/?A|At Loss)).*?"
    r"Current:\s*([+-]?(?:\d+(?:\.\d+)?|N/?A|At Loss))",
    re.I | re.S,
)
MIN_RE = re.compile(r"\bMin:\s*([+-]?\d+(?:\.\d+)?)", re.I)
MED_RE = re.compile(r"\bMed:\s*([+-]?\d+(?:\.\d+)?)", re.I)
MAX_RE = re.compile(r"\bMax:\s*([+-]?\d+(?:\.\d+)?)", re.I)
CUR_RE = re.compile(r"\bCurrent:\s*([+-]?\d+(?:\.\d+)?)", re.I)
# 正文兜底： "is 37.81 as of ... 10-year median of 40.18"
BODY_RE = re.compile(
    r"(?:is|at)\s+([+-]?\d+(?:\.\d+)?)\s+as of.{0,80}?10-year median of\s+([+-]?\d+(?:\.\d+)?)",
    re.I | re.S,
)
TITLE_RE = re.compile(
    r"(?:PE Ratio(?:\s*\(TTM\))?|PB Ratio|PS Ratio|EV-to-FCF|EV/FCF|EV-to-Free Cash Flow)\s*:\s*([+-]?\d+(?:\.\d+)?)",
    re.I,
)
# 禁止用「标题里第一个冒号后的数字」当 Current。
# 美团 PE 页标题是 Meituan (HKSE:03690) …，亏损时 Current=At Loss，
# 这个正则会把 03690 写成 PE=3690。
TITLE_CUR_RE = None
MEDIAN_ONLY_RE = re.compile(
    r"(?:10-year median of|10-Year Median[:\s]+|Median[:\s]+)\s*([+-]?\d+(?:\.\d+)?)",
    re.I,
)
AT_LOSS_RE = re.compile(r"\b(?:Current:\s*)?(?:At Loss|N/?A)\b", re.I)
EXCH_TICKER_RE = re.compile(
    r"\b(?:HKSE|SAU|BSP|ISX|XKLS|NSE|BKK|TPE|TSE|XKRX|XPAR|XSWX|XTER|XMAD|XAMS|LSE|MIL|WBAG|SGX|STC):0*(\d{3,6})\b",
    re.I,
)

BLUE = Font(name="Arial", color="0000FF", size=10)
NOTE_FONT = Font(name="Arial", size=9, color="000000")

HK_RE = re.compile(r"^0*(\d{3,5})\.HK$", re.I)
TW_RE = re.compile(r"^(\d{4})\.TW$", re.I)
KS_RE = re.compile(r"^(\d{6})\.KS$", re.I)
JP_RE = re.compile(r"^(\d{4})\.T$", re.I)
BK_RE = re.compile(r"^([A-Z0-9]+)\.BK$", re.I)
NS_RE = re.compile(r"^([A-Z0-9]+)\.NS$", re.I)
JK_RE = re.compile(r"^([A-Z0-9]+)\.JK$", re.I)
KL_RE = re.compile(r"^(\d{4})\.KL$", re.I)
# 巴西 B3：VALE3 / PETR4 / BPAC11。沙特 Tadawul 四位数字不自动猜（和年份/别国代码撞），必须走对照表 SAU:
BR_RE = re.compile(r"^[A-Z]{4}\d{1,2}$")


def parse_num(s: str | None):
    if s is None:
        return None
    t = str(s).strip().replace(",", "")
    if t == "" or t.upper() in {"N/A", "NA", "AT LOSS", "-", "NONE"}:
        return None
    try:
        return float(t)
    except ValueError:
        return None


def load_map(path: Path) -> dict:
    out = {}
    if not path.exists():
        return out
    with path.open("r", encoding="utf-8-sig") as f:
        for row in csv.DictReader(
            (ln for ln in f if ln.strip() and not ln.lstrip().startswith("#"))
        ):
            local = (row.get("local_code") or "").strip()
            if not local:
                continue
            out[local.upper()] = {
                "gf": (row.get("gf_ticker") or "").strip(),
                "no_pe": (row.get("no_pe") or "").strip() in {"1", "Y", "y", "TRUE", "true"},
                "no_pb": (row.get("no_pb") or "").strip() in {"1", "Y", "y", "TRUE", "true"},
                "no_ps": (row.get("no_ps") or "").strip() in {"1", "Y", "y", "TRUE", "true"},
                "no_ev": (row.get("no_ev") or "").strip() in {"1", "Y", "y", "TRUE", "true"},
                "skip": (row.get("skip") or "").strip() in {"1", "Y", "y", "TRUE", "true"},
                "note": (row.get("note") or "").strip(),
            }
    return out


def candidates(local: str, mapping: dict) -> list[str]:
    raw = str(local).strip()
    key = raw.upper()
    meta = mapping.get(key, {})
    ordered = []
    if meta.get("gf"):
        ordered.append(meta["gf"])
    m = HK_RE.match(raw)
    if m:
        digits = m.group(1)
        ordered += [f"HKSE:{digits.zfill(5)}", f"HKSE:{digits.zfill(4)}", f"HKSE:{digits}"]
    m = TW_RE.match(raw)
    if m:
        ordered += [f"TPE:{m.group(1)}"]
    m = KS_RE.match(raw)
    if m:
        ordered += [f"XKRX:{m.group(1)}"]
    m = JP_RE.match(raw)
    if m:
        ordered += [f"TSE:{m.group(1)}"]
    m = BK_RE.match(raw)
    if m:
        ordered += [f"BKK:{m.group(1)}"]
    m = NS_RE.match(raw)
    if m:
        ordered += [f"NSE:{m.group(1)}"]
    m = JK_RE.match(raw)
    if m:
        ordered += [f"ISX:{m.group(1)}", f"IDX:{m.group(1)}"]
    m = KL_RE.match(raw)
    if m:
        ordered += [f"XKLS:{m.group(1)}", f"KLSE:{m.group(1)}"]
    if BR_RE.match(key) and not meta.get("gf"):
        ordered += [f"BSP:{key}"]
    ordered.append(raw)
    # de-dup keep order
    seen = set()
    out = []
    for x in ordered:
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return out


def extract_range(text: str, title: str = ""):
    """从 term 页正文或标题抽出 current + 十年 med。"""
    t = (
        (text or "")
        .replace("\xa0", " ")
        .replace("&nbsp;", " ")
        .replace("\u2009", " ")
    )
    out = {"min": None, "med": None, "max": None, "current": None}

    m = RANGE_RE.search(t)
    if m:
        out["min"] = parse_num(m.group(1))
        out["med"] = parse_num(m.group(2))
        out["max"] = parse_num(m.group(3))
        out["current"] = parse_num(m.group(4))

    if out["med"] is None:
        mm = MED_RE.search(t)
        if mm:
            out["med"] = parse_num(mm.group(1))
    if out["current"] is None:
        cm = CUR_RE.search(t)
        if cm:
            out["current"] = parse_num(cm.group(1))
    if out["min"] is None:
        mn = MIN_RE.search(t)
        if mn:
            out["min"] = parse_num(mn.group(1))
    if out["max"] is None:
        mx = MAX_RE.search(t)
        if mx:
            out["max"] = parse_num(mx.group(1))

    if out["current"] is None or out["med"] is None:
        bm = BODY_RE.search(t)
        if bm:
            if out["current"] is None:
                out["current"] = parse_num(bm.group(1))
            if out["med"] is None:
                out["med"] = parse_num(bm.group(2))
    if out["med"] is None:
        mo = MEDIAN_ONLY_RE.search(t)
        if mo:
            out["med"] = parse_num(mo.group(1))
    # 标题只认「PE Ratio: 12.3」这种，不认 HKSE:03690
    if out["current"] is None and title:
        tm = TITLE_RE.search(title)
        if tm:
            out["current"] = parse_num(tm.group(1))

    blob = f"{title}\n{t}"
    if out["current"] is None and AT_LOSS_RE.search(blob):
        out["current"] = None  # 亏损：保持空，不要用代码数字顶上

    if out["current"] is None and out["med"] is None:
        return None
    return out


def ticker_code_numbers(ticker: str) -> set[float]:
    """HKSE:03690 / 3690.HK / 2222 → {3690, 3690.0, 2222, …}"""
    out: set[float] = set()
    s = str(ticker or "")
    for m in re.findall(r"\d{3,6}", s):
        try:
            out.add(float(m))
            out.add(float(m.lstrip("0") or "0"))
        except ValueError:
            pass
    for m in EXCH_TICKER_RE.finditer(s):
        try:
            out.add(float(m.group(1)))
        except ValueError:
            pass
    return {x for x in out if x > 0}


def looks_like_ticker_code(value, ticker: str) -> bool:
    if value is None:
        return False
    try:
        v = float(value)
    except (TypeError, ValueError):
        return False
    if v != int(v):
        return False
    return v in ticker_code_numbers(ticker)


def sanitize_parsed(parsed: dict | None, ticker: str) -> dict | None:
    """丢掉把交易所代码写进倍数的结果。"""
    if not parsed:
        return parsed
    out = dict(parsed)
    for key in ("current", "med", "min", "max"):
        if looks_like_ticker_code(out.get(key), ticker):
            out[key] = None
    if out.get("current") is None and out.get("med") is None:
        return None
    return out


PROFILE_DIR = HERE / "gf_chrome_profile"


def is_missing_symbol(url: str, title: str, text: str) -> bool:
    blob = f"{url}\n{title}\n{text}".lower()
    return (
        "/search?" in blob
        or "results similar to" in blob
        or "did not match any" in blob
        or "no results for" in blob
        or "symbol not found" in blob
    )


def is_hard_block(title: str, text: str) -> bool:
    """真拦截：Sorry you have been blocked / Attention Required。见到就停。"""
    raw = f"{title}\n{text}"
    blob = raw.lower()
    return (
        "sorry, you have been blocked" in blob
        or "attention required" in blob
        or ("cloudflare ray id" in blob and "blocked" in blob)
        or "您已被阻止" in raw
        or "已被封锁" in raw
    )


def is_cf_challenge(title: str, text: str) -> bool:
    """验证中间页：Just a moment / 正在进行安全验证。通常几秒到几十秒自己过，或手点一下过。不算拦截。"""
    if is_hard_block(title, text):
        return False
    raw = f"{title}\n{text}"
    blob = raw.lower()
    return (
        "just a moment" in blob
        or "checking your browser" in blob
        or "verify you are human" in blob
        or "正在进行安全验证" in raw
        or "安全验证" in raw
        or "恶意自动程序" in raw
    )


def is_cloudflare_block(title: str, text: str) -> bool:
    """旧接口：真拦截或验证页都算（仅无头 goto 模式使用，那里自带等待）。"""
    return is_hard_block(title, text) or is_cf_challenge(title, text)


class GfClient:
    def __init__(
        self,
        cache: dict,
        delay: float,
        headed: bool = False,
        debug: bool = False,
        cdp: str | None = None,
    ):
        self.cache = cache
        self.delay = delay
        self.headed = headed
        self.debug = debug
        self.cdp = cdp
        self.blocked = False
        self.stalled = False  # 卡在验证页没过：停开新页，但不算拦截
        self._pw = None
        self._browser = None
        self._context = None
        self._page = None

    def start(self):
        from playwright.sync_api import sync_playwright

        self._pw = sync_playwright().start()
        if self.cdp:
            print(f"连接已打开的 Chrome：{self.cdp}")
            try:
                self._browser = self._pw.chromium.connect_over_cdp(self.cdp)
            except Exception:
                print("9222 还没开。本轮会按页换新隐身窗口，不先开空窗口。")
                return
            self._context = self._browser.contexts[0] if self._browser.contexts else self._browser.new_context()
            self._page = self._context.pages[0] if self._context.pages else self._context.new_page()
            print(f"已连上调试 Chrome，当前标签 {len(self.iter_pages())} 个")
            return

        PROFILE_DIR.mkdir(exist_ok=True)
        launch_kw = dict(
            user_data_dir=str(PROFILE_DIR),
            headless=not self.headed,
            locale="en-US",
            viewport={"width": 1400, "height": 900},
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        # 优先用本机安装的 Google Chrome，比 Playwright 自带 Chromium 少被墙
        try:
            self._context = self._pw.chromium.launch_persistent_context(
                channel="chrome", **launch_kw
            )
            print("使用本机 Google Chrome + 本地配置目录 gf_chrome_profile")
        except Exception as e:
            print(f"本机 Chrome 不可用（{e}），改用 Playwright Chromium")
            self._context = self._pw.chromium.launch_persistent_context(**launch_kw)
        self._context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        self._page = self._context.pages[0] if self._context.pages else self._context.new_page()

    def reconnect(self):
        if self._browser:
            try:
                self._browser.close()
            except Exception:
                pass
        self._browser = self._pw.chromium.connect_over_cdp(self.cdp or "http://127.0.0.1:9222")
        self._context = self._browser.contexts[0] if self._browser.contexts else self._browser.new_context()
        self._page = self._context.pages[0] if self._context.pages else self._context.new_page()

    def close(self):
        try:
            if self.cdp:
                return
            if self._context:
                self._context.close()
            elif self._browser:
                self._browser.close()
        finally:
            if self._pw and not self.cdp:
                self._pw.stop()
            elif self._pw and self.cdp:
                self._pw.stop()

    def _dump(self, ticker: str, slug: str, title: str, text: str):
        path = HERE / "gf_debug.txt"
        snippet = (title or "") + "\n\n" + (text or "")[:4000]
        path.write_text(snippet, encoding="utf-8")
        print(f"    (已把页面原文前4000字写到 {path.name})")

    def _save_cache(self):
        try:
            CACHE_JSON.write_text(
                json.dumps(self.cache, ensure_ascii=False, indent=1),
                encoding="utf-8",
            )
        except Exception:
            pass

    def iter_pages(self):
        seen = []
        pages = []
        if self._context:
            pages.extend(self._context.pages)
        if self._browser:
            for ctx in self._browser.contexts:
                pages.extend(ctx.pages)
        out = []
        for p in pages:
            if id(p) not in seen:
                seen.append(id(p))
                out.append(p)
        return out

    def classify_term(self, url: str, title: str = "") -> str | None:
        low = f"{url} {title}".lower()
        if "pettm" in low or "/term/pe/" in low:
            return "pe"
        if "pb-ratio" in low or "/term/pb/" in low:
            return "pb"
        if "ps-ratio" in low or "/term/ps/" in low:
            return "ps"
        if "enterprise-value-to-fcf" in low or "ev-to-fcf" in low or "ev2fcf" in low:
            return "ev"
        return None

    def open_term_tabs(self, ticker: str, specs) -> list:
        """在已连接的 Chrome 里新开 4 个 term 标签，不关其它窗口。"""
        opened = []
        ctx = self._context
        if ctx is None and self._browser and self._browser.contexts:
            ctx = self._browser.contexts[0]
        if ctx is None:
            print("    没有可开标签的 Chrome 上下文，先确认 --cdp 已连上。")
            return opened
        for field, _c, _m, _i, slugs in specs:
            url = f"https://www.gurufocus.com/term/{slugs[0]}/{ticker}"
            try:
                page = ctx.new_page()
                page.goto(url, wait_until="domcontentloaded", timeout=45000)
                opened.append(page)
                print(f"    已打开 {field.upper()}  {url}")
                time.sleep(1.2)
            except Exception as e:
                print(f"    打开 {field} 失败: {e}")
        return opened

    def close_term_tabs(self, ticker: str) -> int:
        n = 0
        tk = ticker.lower()
        for page in list(self.iter_pages()):
            try:
                url = (page.url or "").lower()
            except Exception:
                continue
            if "gurufocus.com/term/" in url and tk in url:
                try:
                    page.close()
                    n += 1
                except Exception:
                    pass
        return n

    def dismiss_popups(self):
        for page in self.iter_pages():
            try:
                if "gurufocus.com" not in (page.url or "").lower():
                    continue
                title = (page.title() or "").lower()
                # 验证页上不要按 Escape，免得打断 Cloudflare 验证
                if "just a moment" in title or "安全验证" in title or "moment" in title:
                    continue
                page.keyboard.press("Escape")
            except Exception:
                pass

    def scan_open_terms(self, ticker: str) -> dict:
        """一次扫调试窗口里所有标签，按 URL 归到 pe/pb/ps/ev。"""
        self.dismiss_popups()
        found = {}
        tk = ticker.lower()
        for page in self.iter_pages():
            try:
                url = page.url or ""
                title = page.title() or ""
            except Exception:
                continue
            if "gurufocus.com" not in url.lower():
                continue
            if tk not in url.lower() and tk not in title.lower():
                continue
            field = self.classify_term(url, title)
            if not field or field in found:
                continue
            try:
                text = page.inner_text("body")
                html = page.content()
            except Exception:
                continue
            parsed = sanitize_parsed(extract_range(text, title) or extract_range(html, title), ticker)
            if parsed and (parsed.get("current") is not None or parsed.get("med") is not None):
                parsed["url"] = url
                parsed["ticker"] = ticker
                found[field] = parsed
        return found

    def list_tab_urls(self) -> list:
        out = []
        for page in self.iter_pages():
            try:
                out.append(page.url or "")
            except Exception:
                continue
        return out

    def read_from_open_tabs(self, ticker: str, slugs: list[str]):
        """只读已打开的标签，绝不 goto。"""
        tk = ticker.lower()
        gf_pages = []
        for page in self.iter_pages():
            try:
                url = page.url or ""
                title = page.title() or ""
            except Exception:
                continue
            low = url.lower()
            if "gurufocus.com" not in low:
                continue
            gf_pages.append((page, url, title, low))

        def try_parse(page, url, title):
            text = ""
            html = ""
            err = ""
            try:
                text = page.inner_text("body")
            except Exception as e:
                err = f"inner_text: {e}"
            try:
                html = page.content()
            except Exception as e:
                err = (err + " | " if err else "") + f"content: {e}"
            parsed = sanitize_parsed(extract_range(text, title) or extract_range(html, title), ticker)
            if parsed and (parsed["current"] is not None or parsed["med"] is not None):
                parsed["url"] = url
                parsed["ticker"] = ticker
                return parsed
            print(f"    读到了标签但没抽出数字。标题: {title!r}")
            if err:
                print(f"    读取错误: {err}")
            snippet = (text or html or "").replace("\n", " ")[:240]
            print(f"    正文开头: {snippet!r}")
            return None

        for page, url, title, low in gf_pages:
            if tk not in low and tk not in title.lower():
                continue
            slug_hit = any(s.lower() in low or s.lower() in title.lower() for s in slugs)
            ev_hit = ("fcf" in low or "fcf" in title.lower()) and (
                slugs[0].startswith("enterprise") or slugs[0].startswith("ev")
            )
            if slug_hit or ev_hit:
                parsed = try_parse(page, url, title)
                if parsed:
                    return parsed
        # 禁止「只有一个 GF 标签就把 PE 页当成 PB/PS/EV」。必须 URL 对得上 slug。
        return None

    def fetch_term(self, ticker: str, slugs: list[str], read_tab: bool = False) -> dict | None:
        cache_key = f"{ticker}|{slugs[0]}"
        raw_hit = cache_lookup(self.cache, ticker, slugs[0])
        if raw_hit:
            hit = sanitize_parsed(raw_hit, ticker)
            if hit and looks_like_ticker_code(raw_hit.get("current"), ticker):
                print(f"    丢弃缓存脏数 {slugs[0]} cur={raw_hit.get('current')}（像代码）")
                for t in ticker_aliases(ticker):
                    self.cache.pop(f"{t}|{slugs[0]}", None)
                hit = None
            if hit:
                print(f"    缓存 {slugs[0]}: cur={hit.get('current')} med={hit.get('med')}")
                return hit
        if self.blocked and not read_tab:
            return None
        if self.cdp and not read_tab:
            parsed = self.read_from_open_tabs(ticker, slugs)
            if parsed:
                cache_store(self.cache, ticker, slugs[0], parsed)
                self._save_cache()
                return parsed
            return None
        if read_tab:
            url = f"https://www.gurufocus.com/term/{slugs[0]}/{ticker}"
            print(f"    请在「调试那只」Chrome（第1步弹出来的）打开：\n      {url}")
            while True:
                try:
                    ans = input("    看到 Min/Med 后回车读取；s 跳过，q 结束：").strip().lower()
                except EOFError:
                    ans = ""
                if ans == "q":
                    self.blocked = True
                    return None
                if ans == "s":
                    return None
                parsed = self.read_from_open_tabs(ticker, slugs)
                if parsed:
                    cache_store(self.cache, ticker, slugs[0], parsed)
                    self._save_cache()
                    print(f"    已读 {slugs[0]}: cur={parsed.get('current')} med={parsed.get('med')}")
                    return parsed
                tabs = self.list_tab_urls()
                print("    没找到这一页。脚本现在能看见的标签：")
                if not tabs:
                    print("      （一个都没有 —— 链接开在日常 Chrome 里了，要开在调试那只里）")
                else:
                    for u in tabs:
                        print(f"      {u}")
                print("    改到调试窗口打开后，再按一次回车。")
        last_err = None
        last_dump = None
        for slug in slugs:
            url = f"https://www.gurufocus.com/term/{slug}/{ticker}"
            try:
                time.sleep(self.delay + random.uniform(0.2, 0.8))
                self._page.goto(url, wait_until="domcontentloaded", timeout=45000)
                # 等正文或 Min/Med 渲染出来；GF 这行经常晚于首屏
                for _ in range(8):
                    title = self._page.title() or ""
                    text = self._page.inner_text("body")
                    blob = title + "\n" + text
                    if "Just a moment" in blob or "Checking your browser" in blob:
                        self._page.wait_for_timeout(1500)
                        continue
                    if (
                        "10-year median" in blob.lower()
                        or "Min:" in blob
                        or "Med:" in blob
                    ):
                        break
                    self._page.wait_for_timeout(700)
                title = self._page.title() or ""
                text = self._page.inner_text("body")
                if is_cloudflare_block(title, text) or "Just a moment" in (title + text):
                    if self.headed:
                        print(
                            "    弹出窗口被 Cloudflare 拦住了。请在窗口里完成验证，"
                            "不要关浏览器，最多等 90 秒…"
                        )
                        cleared = False
                        for _ in range(30):
                            self._page.wait_for_timeout(3000)
                            title = self._page.title() or ""
                            text = self._page.inner_text("body")
                            if not is_cloudflare_block(title, text) and "Just a moment" not in (title + text):
                                cleared = True
                                break
                        if not cleared:
                            self.blocked = True
                            print("    90 秒内没过验证，停止后续抓取，以免继续被封。")
                            if self.debug:
                                self._dump(ticker, slug, title, text)
                            return None
                    else:
                        self.blocked = True
                        print(
                            "    Cloudflare 已拦截自动化访问。请加 --headed 在窗口里点验证，"
                            "或用 --cdp 连接已打开的 Chrome。"
                        )
                        if self.debug:
                            self._dump(ticker, slug, title, text)
                        return None
                parsed = sanitize_parsed(extract_range(text, title), ticker)
                if parsed and (parsed["current"] is not None or parsed["med"] is not None):
                    parsed["url"] = url
                    parsed["ticker"] = ticker
                    cache_store(self.cache, ticker, slugs[0], parsed)
                    self._save_cache()
                    return parsed
                last_err = f"no-range {url} title={title[:80]!r}"
                last_dump = (ticker, slug, title, text)
            except Exception as e:
                last_err = f"{type(e).__name__}: {e}"
        # 失败不写入 cache，方便下次重试
        if last_err:
            print(f"    ! {ticker} {slugs[0]} 失败: {last_err}")
        if self.debug and last_dump:
            self._dump(*last_dump)
        return None


def iter_holding_rows(ws):
    for r in range(5, ws.max_row + 1):
        name = ws.cell(r, 1).value
        code = ws.cell(r, 2).value
        w = ws.cell(r, 3).value
        if code is None and name is None:
            break
        if code is None:
            continue
        try:
            wf = float(w) if w is not None else None
        except (TypeError, ValueError):
            wf = None
        yield r, str(name or ""), str(code).strip(), wf


def set_blue(cell, value):
    cell.value = value
    cell.font = BLUE


def include_flag(current, median, forced_off: bool, field: str = "") -> int:
    if forced_off:
        return 0
    if current is None or median is None:
        return 0
    # EV/FCF：负数也进（当一家公司、加权、调和都计入）。只看银行/券商的 no_ev。
    # PE/PB/PS：现价或中位 ≤ 0 不进。
    if field == "ev":
        if median <= 0 or current == 0:
            return 0
        return 1
    if current <= 0 or median <= 0:
        return 0
    return 1


def collect_sheet_urls(ws, mapping, min_weight, limit, cache=None, refresh=False):
    urls = []
    seen = set()
    n = 0
    skipped = 0
    for r, name, code, w in iter_holding_rows(ws):
        if min_weight is not None and (w is None or w < min_weight):
            continue
        meta = mapping.get(str(code).upper(), {})
        if meta.get("skip"):
            continue
        n += 1
        if limit and n > limit:
            break
        ticks = candidates(str(code), mapping)
        gf = ticks[0]
        for field, c_cur, c_med, _, slugs in TERM_SPECS:
            # 银行/券商/保险等 no_ps / no_ev = 1：不抓，不提示开页
            if meta.get(f"no_{field}"):
                continue
            if not refresh:
                hit = None
                for t in ticks + [gf]:
                    hit = cache_lookup(cache, t, slugs[0])
                    if hit:
                        break
                # 半个月内取过就跳过，跨页共用。现价是 None 也跳过，不再反复开。
                if cache_usable(hit):
                    skipped += 1
                    continue
            url = f"https://www.gurufocus.com/term/{slugs[0]}/{gf}"
            if url not in seen:
                seen.add(url)
                urls.append(url)
    print(f"    待开 {len(urls)} 个 term，半个月内缓存跳过 {skipped} 个（含现价为 None）")
    return urls


def sheet_term_coverage(ws, mapping, cache, min_weight=0.01):
    need = got = 0
    for r, name, code, w in iter_holding_rows(ws):
        if min_weight is not None and (w is None or w < min_weight):
            continue
        meta = mapping.get(str(code).upper(), {})
        if meta.get("skip"):
            continue
        ticks = candidates(str(code), mapping)
        for field, _c, _m, _i, slugs in TERM_SPECS:
            if meta.get(f"no_{field}"):
                continue
            need += 1
            if any(cache_usable(cache_lookup(cache, t, slugs[0])) for t in ticks):
                got += 1
    return got, need


def flush_cache_to_sheet(ws, mapping, cache, min_weight=0.01) -> int:
    today = date.today().isoformat()
    wrote = 0
    for r, name, code, w in iter_holding_rows(ws):
        if min_weight is not None and (w is None or w < min_weight):
            continue
        meta = mapping.get(str(code).upper(), {})
        if meta.get("skip"):
            continue
        ticks = candidates(str(code), mapping)
        used = None
        any_hit = False
        for field, c_cur, c_med, c_inc, slugs in TERM_SPECS:
            hit = None
            for t in ticks:
                hit = cache_lookup(cache, t, slugs[0])
                if cache_usable(hit):
                    used = t
                    break
                hit = None
            if not hit:
                continue
            any_hit = True
            cur, med = hit.get("current"), hit.get("med")
            if cur is not None:
                set_blue(ws.cell(r, c_cur), cur)
            if med is not None:
                set_blue(ws.cell(r, c_med), med)
            ws.cell(r, c_inc).value = include_flag(cur, med, bool(meta.get(f"no_{field}")), field)
        if any_hit:
            bits = [f"GF自动 {today}"]
            for field, c_cur, c_med, _c_inc, slugs in TERM_SPECS:
                if meta.get(f"no_{field}"):
                    bits.append(f"{field}口径不进")
                    continue
                hit = None
                for t in ticks:
                    hit = cache_lookup(cache, t, slugs[0])
                    if cache_usable(hit):
                        break
                    hit = None
                if not hit:
                    bits.append(f"{field}未取到")
                elif hit.get("current") in (None, ""):
                    bits.append(f"{field}现价None")
                else:
                    bits.append(f"{field}取到")
            set_row_note(ws, r, bits, meta.get("note") or "")
            wrote += 1
    return wrote


def maybe_write_ready_sheets(wb, xlsx, targets, mapping, cache, min_weight, dry_run, snapshot: bool = True):
    """达到 80% 覆盖就写回蓝字。snapshot=True 时才追加当一家公司历史（一次更新只在最后一次调用时做）。"""
    ready = []
    for s in targets:
        got, need = sheet_term_coverage(wb[s], mapping, cache, min_weight)
        pct = (got / need) if need else 0
        print(f"    {s} 覆盖 {got}/{need} = {pct:.0%}")
        if need and pct >= 0.8:
            ready.append((s, int(round(pct * 100))))
    if not ready:
        print("还没有页达到 80%，先不写 Excel。")
        return
    if dry_run:
        print("达到 80%：" + "、".join(s for s, _ in ready) + "。干跑，不写。")
        return
    stamp = _now_stamp()
    for s, cov in ready:
        n = flush_cache_to_sheet(wb[s], mapping, cache, min_weight)
        if snapshot:
            snapshot_as_one_company(wb, wb[s], stamp, source="GF自动", cov_pct=cov)
            print(f"    已写入 {s} {n} 行，并记历史（覆盖 {cov}%）。")
        else:
            print(f"    已写入 {s} {n} 行（本批不追加历史，等最后一次）。")
    refresh_overview_rank(wb)
    wb.save(xlsx)
    print(f"已写回 {xlsx}")


def _term_key_from_url(url: str, title: str = ""):
    field = None
    try:
        # classify without client
        low = f"{url} {title}".lower()
        if "pettm" in low or "/term/pe/" in low:
            field = "pe"
        elif "pb-ratio" in low or "/term/pb/" in low:
            field = "pb"
        elif "ps-ratio" in low or "/term/ps/" in low:
            field = "ps"
        elif "enterprise-value-to-fcf" in low or "ev-to-fcf" in low or "ev2fcf" in low:
            field = "ev"
    except Exception:
        field = None
    ticker = norm_ticker(url.rstrip("/").split("/")[-1] if url else "")
    slug = {
        "pe": "pettm",
        "pb": "pb-ratio",
        "ps": "ps-ratio",
        "ev": "enterprise-value-to-fcf",
    }.get(field)
    return ticker, field, slug


def ingest_all_tabs(client, only_urls: list[str] | None = None):
    """把调试 Chrome 里 GF term 标签读进缓存。"""
    want = None
    if only_urls:
        want = {u.lower().rstrip("/") for u in only_urls}
    n = 0
    blocked = 0
    client.dismiss_popups()
    for page in client.iter_pages():
        try:
            url = page.url or ""
            title = page.title() or ""
        except Exception:
            continue
        if "gurufocus.com/term/" not in url.lower():
            continue
        if want and url.lower().rstrip("/") not in want:
            # 允许重定向后路径微调
            if not any(url.lower().rstrip("/").endswith(u.split("/")[-1].lower()) and
                       any(s in url.lower() for s in ("pettm", "pb-ratio", "ps-ratio", "enterprise-value", "ev-to-fcf"))
                       for u in only_urls):
                continue
        if is_cloudflare_block(title, page.inner_text("body") if False else title):
            pass
        ticker, field, slug = _term_key_from_url(url, title)
        if not field or not ticker or not slug:
            continue
        try:
            text = page.inner_text("body")
            html = page.content()
        except Exception:
            continue
        if is_cloudflare_block(title, text):
            blocked += 1
            continue
        parsed = sanitize_parsed(extract_range(text, title) or extract_range(html, title), ticker)
        if not parsed or parsed.get("med") is None:
            continue
        parsed["url"] = url
        parsed["ticker"] = ticker
        cache_store(client.cache, ticker, slug, parsed)
        n += 1
        print(f"    读到 {ticker} {field}: cur={parsed.get('current')} med={parsed.get('med')}")
    client._save_cache()
    return n, blocked


def _chunk_id(url: str) -> tuple[str, str] | None:
    ticker, field, slug = _term_key_from_url(url, "")
    if not ticker or not slug:
        return None
    return ticker.upper(), slug.lower()


CHALLENGE_GRACE = 30  # 验证页首次出现后至少再等这么多秒


def wait_for_medians(client, urls: list[str], timeout: int = 45):
    """只统计本批 URL，每秒扫一次直到中位出现。
    返回 (读到数, 真拦截数, 仍卡在验证页数)。
    验证页（Just a moment / 正在进行安全验证）不算拦截，首次出现后至少再等 CHALLENGE_GRACE 秒。
    """
    expect = {_chunk_id(u) for u in urls}
    expect.discard(None)
    got = {}
    blocked = set()
    missing = set()
    challenge_first: dict = {}
    deadline = timeout
    sec = 0
    challenging_now: set = set()
    while sec < deadline:
        client.dismiss_popups()
        new_hits = 0
        challenging_now = set()
        for page in client.iter_pages():
            try:
                url = page.url or ""
                title = page.title() or ""
                text = page.inner_text("body")
                html = page.content()
            except Exception:
                continue
            if is_missing_symbol(url, title, text):
                low = url.lower().replace("%3a", ":")
                for ident in list(expect):
                    if ident[0].lower() in low:
                        if ident not in missing:
                            print(f"    代码在GF上不存在（搜索无结果）: {ident[0]} {url}")
                        missing.add(ident)
                continue
            ident = _chunk_id(url)
            if ident is None:
                continue
            if ident not in expect:
                matched = None
                for et, es in list(expect):
                    if ident[1] != es:
                        continue
                    aliases = set(x.upper() for x in ticker_aliases(ident[0]) + ticker_aliases(et))
                    if ident[0].upper() in aliases and et.upper() in aliases:
                        matched = (et, es)
                        break
                if not matched:
                    continue
                ident = matched
            if is_hard_block(title, text):
                if ident not in blocked:
                    print(f"    真拦截页（Sorry you have been blocked / Attention Required）: {url}")
                    blocked.add(ident)
                continue
            if is_cf_challenge(title, text):
                challenging_now.add(ident)
                if ident not in challenge_first:
                    challenge_first[ident] = sec
                    new_deadline = max(deadline, sec + CHALLENGE_GRACE)
                    print(
                        f"    {sec+1}s 验证页（不算拦截），至少再等 {CHALLENGE_GRACE} 秒看能否自己通过，"
                        f"也可在窗口里手点：{url}"
                    )
                    deadline = new_deadline
                continue
            ticker, field, slug = _term_key_from_url(url, title)
            parsed = sanitize_parsed(extract_range(text, title) or extract_range(html, title), ticker or "")
            if not parsed or parsed.get("med") is None:
                continue
            key = f"{ticker}|{slug}"
            if key not in got:
                parsed["url"] = url
                parsed["ticker"] = ticker
                cache_store(client.cache, ticker, slug, parsed)
                got[key] = parsed
                new_hits += 1
                note = "（验证已通过）" if ident in challenge_first else ""
                print(f"    {sec+1}s 读到 {ticker} {field}: cur={parsed.get('current')} med={parsed.get('med')}{note}")
        if new_hits:
            client._save_cache()
        if blocked:
            print(f"    本页真拦截（读到 {len(got)}，拦截 {len(blocked)}），先记一笔，不立刻停整轮")
            break
        if len(got) >= len(expect):
            print(f"    本批 {len(got)}/{len(expect)} 已齐，用了 {sec+1} 秒")
            client._save_cache()
            return len(got), 0, 0
        done = len(got) + len(missing)
        if done >= len(expect) and not challenging_now and sec >= 3:
            print(f"    本批可读 {len(got)}、无此代码 {len(missing)}，不再空等")
            break
        time.sleep(1)
        sec += 1
        if sec in (5, 10, 20, 30, 45, 60) and len(got) < len(expect):
            extra = f"，其中 {len(challenging_now)} 页仍在验证" if challenging_now else ""
            print(f"    已等 {sec}s，本批读到 {len(got)}/{len(expect)}{extra}")
    still = len(challenging_now)
    print(f"    本批结束：读到 {len(got)}/{len(expect)}，真拦截 {len(blocked)}，仍卡验证 {still}")
    client._save_cache()
    return len(got), len(blocked), still


def close_batch_tabs(client, urls: list[str]) -> int:
    """只关本批开的 term 标签；用户自己开的页、窗口本身都不动。"""
    expect = {_chunk_id(u) for u in urls}
    expect.discard(None)
    n = 0
    for page in list(client.iter_pages()):
        try:
            ident = _chunk_id(page.url or "")
        except Exception:
            continue
        if ident and ident in expect:
            try:
                page.close()
                n += 1
            except Exception:
                pass
    if n == 0:
        n = cdp_close_tabs(urls)
    return n


def write_click_html(path: Path, titled_urls: list[tuple[str, str]]) -> None:
    lines = [
        "<!doctype html><meta charset=utf-8><title>待开</title>",
        "<p>在能打开 GuruFocus 的浏览器里 Ctrl+单击。全部打开后再跑 --read-open。</p>",
    ]
    for title, url in titled_urls:
        lines.append(f'<p><a href="{url}" target="_blank">{title}</a></p>')
    path.write_text("\n".join(lines), encoding="utf-8")
    print(f"可 Ctrl+单击的网址已写入 {path} ，共 {len(titled_urls)} 条。脚本不开浏览器。")


def process_sheet(
    ws,
    client,
    mapping,
    min_weight,
    dry_run,
    limit,
    stats,
    read_tab=False,
    open_tabs=False,
    batch_size=4,
    batch_wait=90,
    wb=None,
    refresh=False,
):
    today = date.today().isoformat()
    if client.blocked:
        print("    上一页已被拦截，本页跳过，避免继续开拦截页。")
        return
    if client.stalled:
        print("    上一页卡在验证没过，本页不再开网页（只用缓存里已有的）。")
    preloaded = False
    if read_tab:
        missing_urls = collect_sheet_urls(
            ws, mapping, min_weight, limit, client.cache, refresh=refresh
        )
        print(
            f"缺 {len(missing_urls)} 个。只读已打开的标签，不再开页。"
        )
        if not client.stalled:
            got, nblock = ingest_all_tabs(client)
            print(f"    已读打开的标签 {got} 个，真拦截 {nblock} 个")
        preloaded = True
        if client.stalled:
            print("    没连上日常 Edge，本页停止，不再逐行空转。")
            return
        if client.blocked:
            print("    本页遇到拦截：不写备注、不追加当一家公司快照，本页原样保留。")
            return
    n = 0
    for r, name, code, w in iter_holding_rows(ws):
        if min_weight is not None and (w is None or w < min_weight):
            continue
        meta = mapping.get(code.upper(), {})
        if meta.get("skip"):
            print(f"  skip {code} {name}")
            stats["skip"] += 1
            continue
        n += 1
        if limit and n > limit:
            break
        tickers = candidates(code, mapping)
        print(f"  R{r} {code} -> {tickers[0]}  {name}")
        if read_tab:
            gf = tickers[0]
            missing = []
            for field, _, _, _, slugs in TERM_SPECS:
                if meta.get(f"no_{field}"):
                    continue
                if not any(client.cache.get(f"{t}|{slugs[0]}") for t in tickers):
                    missing.append(field)
            if not missing:
                print("    缓存已齐，跳过开页。")
            else:
                print(f"    还缺 {','.join(missing)}。请按下面提示在调试 Chrome 里亲手打开对应网址。")
        results = {}
        used_ticker = None
        for field, c_cur, c_med, c_inc, slugs in TERM_SPECS:
            if meta.get(f"no_{field}"):
                results[field] = None
                continue
            parsed = None
            for t in tickers:
                parsed = client.fetch_term(t, slugs, read_tab=read_tab and not preloaded)
                if parsed:
                    used_ticker = t
                    break
            results[field] = parsed
            stats["fetch"] += 1
            if parsed:
                stats["ok"] += 1
            else:
                stats["miss"] += 1

        if dry_run:
            for field, *_ in TERM_SPECS:
                p = results[field]
                if p:
                    print(f"     {field}: cur={p['current']} med={p['med']}")
                else:
                    print(f"     {field}: (空)")
            continue

        if not any(results.values()):
            print("    四项都没取到，本行不动。")
            continue
        note_bits = [f"GF自动 {today}"]
        if used_ticker and used_ticker.upper() != code.upper():
            note_bits.append(used_ticker)
        for field, c_cur, c_med, c_inc, slugs in TERM_SPECS:
            p = results[field]
            p = sanitize_parsed(p, used_ticker or code) if p else None
            results[field] = p
            cur = p["current"] if p else None
            med = p["med"] if p else None
            if p is None:
                if ws.cell(r, c_cur).value not in (None, "") and ws.cell(r, c_med).value not in (None, ""):
                    note_bits.append(f"{field}沿用表内")
                else:
                    note_bits.append(f"{field}未取到")
                continue
            if looks_like_ticker_code(cur, used_ticker or code):
                cur = None
            if looks_like_ticker_code(med, used_ticker or code):
                med = None
            if cur is not None:
                set_blue(ws.cell(r, c_cur), cur)
            if med is not None:
                set_blue(ws.cell(r, c_med), med)
            forced = bool(meta.get(f"no_{field}"))
            ws.cell(r, c_inc).value = include_flag(cur, med, forced, field)
            if med is None:
                note_bits.append(f"{field}无十年中位")
        set_row_note(ws, r, note_bits, meta.get("note") or "")
        stats["wrote"] += 1
        stats["wrote_sheet"] = stats.get("wrote_sheet", 0) + 1
    if not dry_run and wb is not None and stats.get("wrote_sheet", 0):
        got, need = sheet_term_coverage(ws, mapping, client.cache, min_weight)
        if need and got / need >= 0.8:
            print(f"    覆盖 {got}/{need}，蓝字已写。历史等本次运行结束再记一条。")
        else:
            print(f"    覆盖 {got}/{need}，不到 80%，先不写历史。")


def _as_one_ratio(ws, col_cur, col_med, col_flag, allow_neg_cur=False):
    num = 0.0  # Σ w/med
    den = 0.0  # Σ w/cur
    used = False
    for r, name, code, w in iter_holding_rows(ws):
        if w is None or w <= 0:
            continue
        cur = ws.cell(r, col_cur).value
        med = ws.cell(r, col_med).value
        flag = ws.cell(r, col_flag).value
        try:
            cur = float(cur)
            med = float(med)
        except (TypeError, ValueError):
            continue
        if med <= 0:
            continue
        # 无论是否允许负现价，都必须「进」=1（银行/券商 no_ev 强制为 0，不进）
        if flag != 1:
            continue
        if allow_neg_cur:
            if cur == 0:
                continue
        else:
            if cur <= 0:
                continue
        num += w / med
        den += w / cur
        used = True
    if not used or den == 0 or (allow_neg_cur and den <= 0):
        return None
    return num / den


def compute_as_one_pack(ws):
    pe = _as_one_ratio(ws, 4, 5, 7, False)
    pb = _as_one_ratio(ws, 8, 9, 11, False)
    ps = _as_one_ratio(ws, 12, 13, 15, False)
    ev = _as_one_ratio(ws, 16, 17, 19, True)
    return pe, pb, ps, ev


def _fmt_ratio(x):
    if x is None:
        return ""
    return round(float(x), 4)


def _now_stamp() -> datetime:
    """历史时间只记日期，不带时分。"""
    d = date.today()
    return datetime(d.year, d.month, d.day)


def ensure_hist_sheet(wb):
    """整本追加总账，一行一次调用、一个估值页。从不覆盖已有行。"""
    if HIST_SHEET in wb.sheetnames:
        ws = wb[HIST_SHEET]
        if not ws.cell(1, 1).value:
            for i, h in enumerate(HIST_COLS, 1):
                cell = ws.cell(1, i)
                cell.value = h
                cell.font = Font(name="Arial", size=10, bold=True)
        return ws
    idx = 1 if "总览对比" in wb.sheetnames else 0
    ws = wb.create_sheet(HIST_SHEET, idx)
    for i, h in enumerate(HIST_COLS, 1):
        cell = ws.cell(1, i)
        cell.value = h
        cell.font = Font(name="Arial", size=10, bold=True)
    widths = {"A": 18, "B": 22, "C": 18, "D": 18, "E": 18, "F": 22, "G": 16}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:G1"
    return ws


def _next_empty_row(ws, start: int, cols: int = 5) -> int:
    last = start - 1
    for r in range(start, (ws.max_row or start) + 2):
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, cols + 1)):
            last = r
        elif r > last + 1 and r > start:
            break
    return last + 1


def _parse_cov(source: str | None) -> int | None:
    """从来源字符串里抽出覆盖百分比，例如 'GF自动 80%' → 80。没有则 None。"""
    if not source:
        return None
    m = re.search(r"(\d{1,3})%", str(source))
    return int(m.group(1)) if m else None


def append_central_history(wb, sheet_name: str, stamp: datetime, pe, pb, ps, ev, source: str, cov_pct: int | None = None) -> bool:
    """写入总账。返回 True 表示覆盖了上一轮不足 100% 的中间行。"""
    ws = ensure_hist_sheet(wb)
    tag = f"{source} {cov_pct}%" if cov_pct is not None else source
    overwritten = False
    target = None
    for r in range(ws.max_row or 1, 1, -1):
        if ws.cell(r, 2).value != sheet_name:
            continue
        prev = _parse_cov(ws.cell(r, 7).value)
        if prev is not None and prev < 100 and (cov_pct is None or cov_pct >= prev):
            target = r
            overwritten = True
        break
    if target is None:
        target = _next_empty_row(ws, 2, cols=7)
    ws.cell(target, 1).value = stamp
    ws.cell(target, 1).number_format = "yyyy-mm-dd"
    ws.cell(target, 2).value = sheet_name
    ws.cell(target, 3).value = _fmt_ratio(pe)
    ws.cell(target, 4).value = _fmt_ratio(pb)
    ws.cell(target, 5).value = _fmt_ratio(ps)
    ws.cell(target, 6).value = _fmt_ratio(ev)
    ws.cell(target, 7).value = tag
    for c in range(1, 8):
        ws.cell(target, c).font = NOTE_FONT
    for c in (3, 4, 5, 6):
        if isinstance(ws.cell(target, c).value, (int, float)):
            ws.cell(target, c).number_format = "0.0000"
    return overwritten


def snapshot_as_one_company(wb, ws, stamp: datetime, source: str = "GF自动", cov_pct: int | None = None):
    """本页底部和总账各记一行。上一轮不足 100% 时覆盖那条，已是 100% 才另起一条。"""
    pe, pb, ps, ev = compute_as_one_pack(ws)
    label_row = None
    hist_header = None
    for r in range(1, (ws.max_row or 1) + 1):
        v = ws.cell(r, 1).value
        if not v:
            continue
        s = str(v).strip()
        if s.startswith("当一家公司：组合PE"):
            label_row = r
        if s == "当一家公司历史":
            hist_header = r
    if label_row:
        ws.cell(label_row, 3).value = f"最近写入 {stamp:%Y-%m-%d}"
        ws.cell(label_row, 3).font = NOTE_FONT

    if hist_header is None:
        start = (ws.max_row or 90) + 2
        ws.cell(start, 1).value = "当一家公司历史"
        ws.cell(start + 1, 1).value = "时间"
        ws.cell(start + 1, 2).value = "组合PE/常态PE"
        ws.cell(start + 1, 3).value = "组合PB/常态PB"
        ws.cell(start + 1, 4).value = "组合PS/常态PS"
        ws.cell(start + 1, 5).value = "组合EVFCF/常态EVFCF"
        hist_header = start
    else:
        # 旧表头「日期」改成「时间」，旧行保留
        hdr = ws.cell(hist_header + 1, 1)
        if hdr.value == "日期":
            hdr.value = "时间"
    data_start = hist_header + 2
    overwritten = append_central_history(wb, ws.title, stamp, pe, pb, ps, ev, source, cov_pct)
    if overwritten:
        # 总账覆盖了中间行，本页底部也覆盖最后一条
        target = None
        for r in range((ws.max_row or data_start), data_start - 1, -1):
            if ws.cell(r, 1).value not in (None, ""):
                target = r
                break
        if target is None:
            target = _next_empty_row(ws, data_start, cols=5)
    else:
        target = _next_empty_row(ws, data_start, cols=5)
    ws.cell(target, 1).value = stamp
    ws.cell(target, 1).number_format = "yyyy-mm-dd"
    ws.cell(target, 2).value = _fmt_ratio(pe)
    ws.cell(target, 3).value = _fmt_ratio(pb)
    ws.cell(target, 4).value = _fmt_ratio(ps)
    ws.cell(target, 5).value = _fmt_ratio(ev)
    for c in range(1, 6):
        ws.cell(target, c).font = NOTE_FONT
    for c in (2, 3, 4, 5):
        if isinstance(ws.cell(target, c).value, (int, float)):
            ws.cell(target, c).number_format = "0.0000"
    print(
        f"    当一家公司快照 {stamp:%Y-%m-%d}  PE={_fmt_ratio(pe)} PB={_fmt_ratio(pb)} "
        f"PS={_fmt_ratio(ps)} EV={_fmt_ratio(ev)}  → 行{target}"
        + ("（覆盖中间行）" if overwritten else "")
    )
    _refresh_history_chart(ws, data_start, target)


def _overview_sheet_of_row(ws, r: int) -> str | None:
    """底部对比行 B 列是 =E45 这类引用，顺着找到源估值页名。"""
    f = ws.cell(r, 2).value
    if not isinstance(f, str) or not f.startswith("="):
        return None
    ref = f[1:].replace("$", "")
    col = "".join(ch for ch in ref if ch.isalpha())
    row_s = "".join(ch for ch in ref if ch.isdigit())
    if not col or not row_s:
        return None
    src = ws.cell(int(row_s), column_index_from_string(col)).value
    if not isinstance(src, str) or "!" not in src:
        return None
    return src.split("!")[0].lstrip("=").strip("'")


def refresh_overview_rank(wb) -> None:
    """总览对比底部表按四维平均重排，并重建图表（避免 openpyxl 存盘后坐标轴消失）。"""
    if "总览对比" not in wb.sheetnames:
        return
    ws = wb["总览对比"]
    # 动态找说明行，扫描到说明行之前（不再写死 73）
    note_row = 73
    for r in range(54, 100):
        v = ws.cell(r, 1).value
        if v and "横坐标已按四维平均" in str(v):
            note_row = r
            break
    block = []
    for r in range(54, note_row):
        name = ws.cell(r, 1).value
        if not name:
            continue
        sheet = _overview_sheet_of_row(ws, r)
        pe = pb = ps = ev = None
        if sheet and sheet in wb.sheetnames:
            pe, pb, ps, ev = compute_as_one_pack(wb[sheet])
        vals = [x for x in (pe, pb, ps, ev) if isinstance(x, (int, float))]
        avg = sum(vals) / len(vals) if vals else None
        cells = [ws.cell(r, c).value for c in range(1, 9)]
        block.append((avg if avg is not None else 9999, cells, pe, pb, ps, ev))
    block.sort(key=lambda x: x[0])
    names = []
    series_vals = [[], [], [], []]
    for i, (_avg, cells, pe, pb, ps, ev) in enumerate(block):
        rr = 54 + i
        for c, v in enumerate(cells, 1):
            if c in (6, 7, 8):
                continue
            ws.cell(rr, c).value = v
        ws.cell(rr, 6).value = 1
        ws.cell(rr, 7).value = 1.5
        ws.cell(rr, 8).value = f'=IF(COUNT(B{rr}:E{rr})=0,"",AVERAGE(B{rr}:E{rr}))'
        for c in (2, 3, 4, 5, 8):
            ws.cell(rr, c).number_format = "0.00"
        names.append(cells[0])
        for bucket, val in zip(series_vals, (pe, pb, ps, ev)):
            bucket.append(val if isinstance(val, (int, float)) else None)
    last = 53 + len(block)
    ws._charts = []
    cat_ref = f"'{ws.title}'!$A$54:$A${last}"
    cats = Reference(ws, min_col=1, min_row=54, max_row=last)
    n = len(names)

    def _fill(hex_color):
        sp = GraphicalProperties()
        sp.solidFill = hex_color
        return sp

    def _dash(hex_color):
        sp = GraphicalProperties()
        sp.ln = LineProperties(solidFill=hex_color, prstDash="dash", w=14000)
        return sp

    # 底色：0–1 浅绿，1–2.5 淡粉（F=1 与 G=1.5 堆叠）
    area = AreaChart()
    area.grouping = "stacked"
    area.title = "当一家公司：相对十年中位（<1 便宜）"
    area.y_axis.delete = False
    area.x_axis.delete = False
    area.y_axis.scaling.min = 0
    area.y_axis.scaling.max = 2.5
    bands = Reference(ws, min_col=6, min_row=53, max_col=7, max_row=last)
    area.add_data(bands, titles_from_data=True)
    area.set_categories(cats)
    for idx, (title, color, const) in enumerate((
        ("便宜区(<1)", "C6EFCE", 1),
        ("贵区(>1)", "F8CBAD", 1.5),
    )):
        s = area.series[idx]
        s.cat = AxDataSource(strRef=StrRef(cat_ref))
        s.title = SeriesLabel(v=title)
        s.graphicalProperties = _fill(color)
        pts = [NumVal(idx=i, v=const) for i in range(n)]
        if s.val and s.val.numRef:
            s.val.numRef.numCache = NumData(ptCount=n, pt=pts)

    # 四维柱
    bar = BarChart()
    bar.type = "col"
    bar.grouping = "clustered"
    bar.y_axis.axId = 200
    data = Reference(ws, min_col=2, min_row=53, max_col=5, max_row=last)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    titles = ["PE/常态", "PB/常态", "PS/常态", "EV/常态"]
    for idx, s in enumerate(bar.series):
        s.cat = AxDataSource(strRef=StrRef(cat_ref))
        s.title = SeriesLabel(v=titles[idx])
        s.graphicalProperties = None
        pts = [NumVal(idx=i, v=v) for i, v in enumerate(series_vals[idx]) if isinstance(v, (int, float))]
        if s.val and s.val.numRef:
            s.val.numRef.numCache = NumData(ptCount=len(pts), pt=pts)
    bar.overlap = -10

    # 1 和 1.5 虚线
    line = LineChart()
    line.y_axis.axId = 200
    line.add_data(bands, titles_from_data=True)
    line.set_categories(cats)
    for idx, (title, color, const) in enumerate((
        ("=1", "548235", 1),
        ("=1.5", "C00000", 1.5),
    )):
        s = line.series[idx]
        s.cat = AxDataSource(strRef=StrRef(cat_ref))
        s.title = SeriesLabel(v=title)
        s.graphicalProperties = _dash(color)
        s.marker = Marker(symbol="none")
        pts = [NumVal(idx=i, v=const) for i in range(n)]
        if s.val and s.val.numRef:
            s.val.numRef.numCache = NumData(ptCount=n, pt=pts)

    area.y_axis.crosses = "min"
    area += bar
    area += line
    area.legend.position = "b"
    area.width = 28
    area.height = 12
    ws.add_chart(area, "A75")
    print(f"    总览对比已按四维平均重排 {len(block)} 行，四维柱状图已重建。")


def _ensure_history_chart(ws, hist_header: int) -> None:
    """如果本页没有散点图，建一个当一家公司历史图。"""
    if ws._charts:
        return
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
    data_start = hist_header + 2
    xvalues = Reference(ws, min_col=1, min_row=data_start, max_row=data_start + 1)
    titles = ["组合PE/常态PE", "组合PB/常态PB", "组合PS/常态PS", "组合EVFCF/常态EVFCF"]
    for col, title in enumerate(titles, 2):
        yvalues = Reference(ws, min_col=col, min_row=hist_header + 1, max_row=data_start + 1)
        s = Series(yvalues, xvalues, title_from_data=True)
        s.title = SeriesLabel(v=title)
        chart.series.append(s)
    # 放在历史区下方
    ws.add_chart(chart, "A{}".format(hist_header + 8))


def _refresh_history_chart(ws, data_start: int, last_row: int) -> None:
    """把本页散点图扩到全部真日期行。跳过备注文字行，避免横轴再变 1900。"""
    # 没有图表就先建一个
    hist_header = data_start - 2
    _ensure_history_chart(ws, hist_header)
    if not ws._charts or last_row < data_start:
        return
    first = None
    last = None
    for r in range(data_start, last_row + 1):
        v = ws.cell(r, 1).value
        if hasattr(v, "year"):
            if first is None:
                first = r
            last = r
    if first is None:
        return
    title = ws.title
    q = f"'{title}'" if any(c in title for c in " !:'") else title
    cols = ["B", "C", "D", "E"]
    for ch in ws._charts:
        for j, s in enumerate(ch.series):
            if j >= len(cols):
                break
            col = cols[j]
            s.xVal = AxDataSource(numRef=NumRef(f="{q}!$A${first}:$A${last}"))
            if s.yVal and s.yVal.numRef:
                s.yVal.numRef.f = f"{q}!${col}${first}:${col}${last}"
                s.yVal.numRef.numCache = None
            if s.xVal and s.xVal.numRef:
                s.xVal.numRef.numCache = None


GF_AUTO_DATE_RE = re.compile(r"GF自动\s+(\d{4}-\d{2}-\d{2})")


def set_row_note(ws, r, bits: list[str], extra: str = "") -> None:
    """只留这一次的结果。旧的 GF自动 日期丢掉，口径备注留一次。"""
    old = str(ws.cell(r, 20).value or "")
    keep = extra
    if not keep and "｜" in old:
        tail = old.split("｜")[-1]
        if "GF自动" not in tail:
            keep = tail
    elif not keep and old and "GF自动" not in old:
        keep = old
    text = "；".join(bits)
    if keep:
        text = text + "｜" + keep[:40]
    ws.cell(r, 20).value = text
    ws.cell(r, 20).font = NOTE_FONT


def sheet_last_update(ws) -> str:
    dates = []
    for r in range(1, min((ws.max_row or 1) + 1, 120)):
        for c in (1, 2, 20):
            v = ws.cell(r, c).value
            if not v:
                continue
            dates.extend(GF_AUTO_DATE_RE.findall(str(v)))
    return max(dates) if dates else "尚无GF自动记录"


def pick_sheets(wb) -> list[str]:
    names = [s for s in wb.sheetnames if s not in SKIP_SHEETS]
    print("\n可更新的估值页：")
    print("   0) 全部")
    for i, name in enumerate(names, 1):
        last = sheet_last_update(wb[name])
        print(f"  {i:2d}) {name}    上次写入：{last}")
    print("输入序号（多个用逗号，直接回车=取消）：")
    try:
        raw = input("> ").strip()
    except EOFError:
        raw = ""
    if not raw:
        sys.exit("未选择页面，退出。")
    if raw == "0":
        return names
    picked = []
    for part in raw.replace("，", ",").split(","):
        part = part.strip()
        if not part:
            continue
        if part.isdigit():
            idx = int(part)
            if 1 <= idx <= len(names):
                picked.append(names[idx - 1])
                continue
        # allow typing sheet name fragment
        hits = [n for n in names if part.lower() in n.lower()]
        if len(hits) == 1:
            picked.append(hits[0])
        elif part in names:
            picked.append(part)
        else:
            sys.exit(f"无法识别：{part}")
    if not picked:
        sys.exit("未选择页面，退出。")
    # de-dup keep order
    out = []
    seen = set()
    for n in picked:
        if n not in seen:
            seen.add(n)
            out.append(n)
    return out


def main():
    ap = argparse.ArgumentParser(description="用 GuruFocus 公开十年中位更新估值底稿")
    ap.add_argument("--file", default=str(DEFAULT_XLSX), help="xlsx 路径")
    ap.add_argument("--sheet", action="append", dest="sheets", help="只更新指定页，可重复")
    ap.add_argument("--all", action="store_true", help="更新除总览/说明外全部估值页")
    ap.add_argument(
        "--snapshot-only",
        action="store_true",
        help="不抓 GuruFocus，只按当前蓝字追加一行「当一家公司」历史",
    )
    ap.add_argument("--min-weight", type=float, default=0.01, help="默认只抓权重>=1%%")
    ap.add_argument("--limit", type=int, default=0, help="每页最多抓几只，0=不限")
    ap.add_argument("--dry-run", action="store_true", help="只打印不写回")
    ap.add_argument("--delay", type=float, default=5.0, help="两次打开网页的间隔秒，默认5秒防封")
    ap.add_argument("--no-cache", action="store_true", help="忽略本地缓存")
    ap.add_argument("--clear-cache", action="store_true", help="删除 gf_cache.json 后退出")
    ap.add_argument("--refresh", action="store_true", help="忽略缓存和表内已有蓝字，强制重抓")
    ap.add_argument("--headed", action="store_true", help="弹出真浏览器窗口，方便看是不是被墙")
    ap.add_argument("--debug", action="store_true", help="失败时把页面原文写到 gf_debug.txt")
    ap.add_argument(
        "--cdp",
        default="",
        help="连接已打开的 Chrome，例如 http://127.0.0.1:9222",
    )
    ap.add_argument(
        "--read-tab",
        action="store_true",
        help="你自己打开网页，脚本只读取当前标签，不会自动跳转",
    )
    ap.add_argument(
        "--open-tabs",
        action="store_true",
        help="每只股票自动打开4个term标签，回车提取后关掉再开下一只",
    )
    ap.add_argument("--batch-size", type=int, default=4, help="每批打开几个term页")
    ap.add_argument("--batch-wait", type=float, default=90, help="两批之间等待秒数，默认90")
    ap.add_argument("--read-open", action="store_true", help="只读你已经打开的标签，不再开浏览器、不再写链接")
    args = ap.parse_args()

    if args.clear_cache:
        if CACHE_JSON.exists():
            CACHE_JSON.unlink()
            print(f"已删除缓存：{CACHE_JSON}")
        else:
            print(f"缓存文件不存在：{CACHE_JSON}")
        return

    xlsx = Path(args.file)
    if not xlsx.exists():
        for cand in (HERE / XLSX_NAME, Path.cwd() / XLSX_NAME, HERE.parent / XLSX_NAME):
            if cand.exists():
                xlsx = cand
                print(f"改用表格：{xlsx}")
                break
    if not xlsx.exists():
        sys.exit(f"找不到文件: {xlsx}。请把表格放到脚本同一文件夹，或加 --file 路径")

    mapping = load_map(MAP_CSV)
    cache = {}
    if CACHE_JSON.exists() and not args.no_cache:
        try:
            raw = json.loads(CACHE_JSON.read_text(encoding="utf-8"))
            cache = {k: v for k, v in raw.items() if v}
        except Exception as e:
            print(f"读取 {CACHE_JSON.name} 失败：{e}，当成空缓存")
            cache = {}
    print(f"缓存 {CACHE_JSON} ：{len(cache)} 条" + ("（本轮强制重抓）" if args.refresh else ""))

    wb = load_workbook(xlsx)
    if args.sheets:
        targets = args.sheets
    elif args.all:
        targets = [s for s in wb.sheetnames if s not in SKIP_SHEETS]
    else:
        print("写表前请先关闭 Excel。只预览请加 --dry-run。")
        targets = pick_sheets(wb)
        args.read_tab = True
        if not args.cdp:
            args.cdp = "http://127.0.0.1:9222"
        print("已选：" + "、".join(targets))
    print("脚本不会自己开 GuruFocus。先写待开.html，你 Ctrl+单击，开完一批按回车读取。")

    if not args.read_open:
        while True:
            urls = []
            for s in targets:
                urls.extend(
                    collect_sheet_urls(
                        wb[s], mapping, args.min_weight, args.limit, cache, refresh=args.refresh
                    )
                )
            batch, seen = [], []
            for u in urls:
                ticker = u.rstrip("/").split("/")[-1]
                if ticker not in seen:
                    if len(seen) >= 5:
                        break
                    seen.append(ticker)
                batch.append(u)
            write_click_html(HERE / "待开.html", [(u, u) for u in batch])
            if not batch:
                print("缓存里都有值了，没有还要开的页。")
                break
            if not _port_open():
                launch_in_daily_edge((HERE / "待开.html").as_uri())
            print(f"本批 {len(seen)} 只：{', '.join(seen)}。只在脚本打开的那只窗口里点。回车不再开新窗口。")
            try:
                cmd = input("开完这一批按回车读取，q 结束：").strip().lower()
            except EOFError:
                cmd = "q"
            if cmd == "q":
                break
            if not _port_open():
                print("没读到，也没有新开浏览器。")
                continue
            client = GfClient(cache, args.delay, cdp="http://127.0.0.1:9222")
            client.start()
            got, nblock = ingest_all_tabs(client)
            print(f"本批读到 {got} 个，真拦截 {nblock} 个。缓存已留着。")
            client.close()
            maybe_write_ready_sheets(
                wb, xlsx, targets, mapping, cache, args.min_weight, args.dry_run, snapshot=False
            )
        maybe_write_ready_sheets(
            wb, xlsx, targets, mapping, cache, args.min_weight, args.dry_run, snapshot=True
        )
        CACHE_JSON.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"已保存缓存 {len(cache)} 条 -> {CACHE_JSON}")
        return

    for s in targets:
        if s not in wb.sheetnames:
            sys.exit(f"没有这一页: {s}  现有: {wb.sheetnames}")

    if args.snapshot_only:
        stamp = _now_stamp()
        print(f"只追加当一家公司历史，不抓网页。时间戳 {stamp}")
        for s in targets:
            print(f"\n===== {s} =====")
            if args.dry_run:
                pe, pb, ps, ev = compute_as_one_pack(wb[s])
                print(
                    f"    将写入 PE={_fmt_ratio(pe)} PB={_fmt_ratio(pb)} "
                    f"PS={_fmt_ratio(ps)} EV={_fmt_ratio(ev)}"
                )
            else:
                snapshot_as_one_company(wb, wb[s], stamp, source="仅快照")
        if not args.dry_run:
            refresh_overview_rank(wb)
            wb.save(xlsx)
            print(f"\n已写回 {xlsx}")
        else:
            print("\n干跑，未写文件。")
        return

    client = GfClient(
        cache,
        args.delay,
        headed=args.headed,
        debug=args.debug,
        cdp=args.cdp or None,
    )
    stats = {"fetch": 0, "ok": 0, "miss": 0, "wrote": 0, "skip": 0}
    try:
        if args.cdp and _port_open():
            print("准备连接已打开的浏览器 …")
        elif args.cdp:
            print("9222 没开。请先按说明用日常 Edge 打开，不要等脚本启动无头 Chrome。")
        else:
            print("启动有界面 Chrome …")
        client.start()
        for s in targets:
            print(f"\n===== {s} =====")
            stats["wrote_sheet"] = 0
            process_sheet(
                wb[s],
                client,
                mapping,
                args.min_weight,
                args.dry_run,
                args.limit,
                stats,
                read_tab=args.read_tab,
                open_tabs=args.open_tabs,
                batch_size=args.batch_size,
                batch_wait=args.batch_wait,
                wb=wb,
                refresh=args.refresh,
            )
            if client.blocked:
                print("已被拦截，后面的页一律不跑。")
                break
        # 一次运行只在最后记历史，80% 和 100% 合成一条
        if not args.dry_run and not client.blocked:
            stamp = _now_stamp()
            for s in targets:
                if s not in wb.sheetnames:
                    continue
                got, need = sheet_term_coverage(wb[s], mapping, cache, args.min_weight)
                if need and got / need >= 0.8:
                    cov = int(round(got / need * 100))
                    snapshot_as_one_company(wb, wb[s], stamp, source="GF自动", cov_pct=cov)
                    print(f"    {s} 历史已记（覆盖 {got}/{need}）")
            refresh_overview_rank(wb)
            wb.save(xlsx)
    except KeyboardInterrupt:
        print("\n中止。正在把已读到的缓存写盘，下次会跳过这些页。")
    finally:
        try:
            client.close()
        except Exception:
            pass
        try:
            CACHE_JSON.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
            print(f"已保存缓存 {len(cache)} 条 -> {CACHE_JSON}")
        except Exception as e:
            print(f"写缓存失败: {e}")

    if args.dry_run:
        print("\n干跑，未写文件。")
    elif client.blocked:
        print("\n本次遇到拦截，Excel 一律不保存（缓存里已读到的数仍保留，下次跳过）。")
    elif stats["wrote"] == 0:
        print("\n没有任何一只取到新数，Excel 不保存。")
    else:
        refresh_overview_rank(wb)
        wb.save(xlsx)
        print(f"\n已写回 {xlsx}")
    print(
        f"抓取 {stats['fetch']} 次 / 成功 {stats['ok']} / 空 {stats['miss']} / "
        f"写回 {stats['wrote']} 只 / 跳过 {stats['skip']}"
    )


if __name__ == "__main__":
    main()


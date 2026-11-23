import datetime
import io
import math
import os
import re
import sys

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties
from playwright.sync_api import sync_playwright

TOKEN = os.environ.get("NOTION_TOKEN", "").strip()
TARGET = (os.environ.get("NOTION_DATABASE_ID", "").strip()
          or os.environ.get("NOTION_DATABASE_URL", "").strip())
UPLOAD = os.environ.get("UPLOAD", "").strip() == "1"
API = "https://api.notion.com/v1"
VERSION = "2025-09-03"
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Notion-Version": VERSION,
    "Content-Type": "application/json",
}
URL = "https://www.chubblife.com.hk/zh/customer-service/fulfillment-ratios-of-dividend.html"
SEARCH = "#fulfillment-section .search-btn.search"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
AUTO = "自動生成"
CUR_NAME = {"HKD": "港元", "USD": "美元"}
CUR_ORDER = ["HKD", "USD"]
NA = "不適用"
PCT = re.compile(r"^\s*(\d+(?:\.\d+)?)\s*%\s*$")
N_YEARS = 11
N_FIXED = 4
HEAD_ROW = 3
FIRST = 4
ROW_H = 24
SEP_H = 8
YEAR_W = 13
LOW_LIMIT = 0.8
HIGH_LIMIT = 1.0
NAVY_HEX = "1F4E79"
NAME_COLOR = "FF1F2933"
CUR_COLOR = "FF7030A0"
SEP_COLOR = "FFE9ECEF"


def solid(argb):
    return PatternFill(fill_type="solid", start_color=argb, end_color=argb)


DARK = solid("FF" + NAVY_HEX)
LIGHT = solid("FFDDEBF7")
PALE = solid("FFF4F7FB")
WHITE = solid("FFFFFFFF")
SEP_FILL = solid(SEP_COLOR)
GOOD_FILL = solid("FFE2F0D9")
LOW_FILL = solid("FFFCE4E4")
WHITE_BOLD = Font(bold=True, color="FFFFFFFF")
BLACK_BOLD = Font(bold=True, color="FF000000")
NAME_FONT = Font(bold=True, color=NAME_COLOR)
CUR_FONT = Font(bold=True, color=CUR_COLOR)
LINK_FONT = Font(bold=True, underline="single", color="FF" + NAVY_HEX)
GREY_FONT = Font(color="FFBFBFBF")
GOOD_FONT = Font(color="FF375623")
LOW_FONT = Font(color="FFC00000")
TITLE_FONT = Font(bold=True, size=16, color="FFFFFFFF")
NOTE_FONT = Font(size=10, color="FF44546A")
MSG_FONT = Font(size=10, italic=True, color="FF595959")
SECTION_FONT = Font(bold=True, size=12, color="FF" + NAVY_HEX)
LINE = Side(style="thin", color="FFD0D7E2")
GROUP = Side(style="thin", color="FF7F9BBF")
BORDER = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
CENTER_FIT = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
LEFT_FIT = Alignment(horizontal="left", vertical="center", shrink_to_fit=True, indent=1)

# ---------- 讀取安達網頁 ----------
JS_DATA = r"""() => {
  const el = document.querySelector('#fulfillment-section');
  const vm = el && el.__vue__;
  if (!vm) return null;
  const src = Array.from(document.querySelectorAll('script[src]')).map(x => x.src)
    .find(x => /hk-zh\.js/.test(x)) || '';
  return JSON.parse(JSON.stringify({
    list: vm.$data.fulfillmentSeriesList, data: vm.$data.data, script: src}));
}"""

JS_SET = r"""(a) => {
  const s = document.querySelectorAll('#fulfillment-section select')[a[0]];
  if (!s) return '找不到選單';
  const i = Array.from(s.options).findIndex(o => o.value === a[1]);
  if (i < 0) return '找不到選項';
  if (s.options[i].disabled) return '選項已停用';
  s.selectedIndex = i;
  s.dispatchEvent(new Event('input', {bubbles: true}));
  s.dispatchEvent(new Event('change', {bubbles: true}));
  return 'ok';
}"""

JS_CLICK = r"""() => {
  const b = document.querySelector('#fulfillment-section .search-btn.search');
  if (b) b.click();
  return !!b;
}"""

JS_RESULT = r"""() => {
  const clean = s => (s || '').replace(/\s+/g, ' ').trim();
  const el = document.querySelector('#fulfillment-section');
  if (!el) return null;
  const d = el.__vue__ ? el.__vue__.$data : {};
  const rows = t => Array.from(t.rows).map(r => Array.from(r.cells).map(x => clean(x.textContent)));
  const out = {series: d.fulfillmentSeries, confirm: d.currencyConfirm,
               desktop: [], mobile: [], intro: '', notes: [], extra: ''};
  const box = el.querySelector('.result-ctn');
  if (!box) return out;
  out.desktop = Array.from(box.querySelectorAll('table.desktop-table')).map(rows);
  out.mobile = Array.from(box.querySelectorAll('table.mobile-table')).map(rows);
  const p = Array.from(box.querySelectorAll('p')).find(x => /日期/.test(x.textContent));
  out.intro = p ? clean(p.textContent) : '';
  const rem = box.querySelector('.remarks-section');
  if (rem) {
    Array.from(rem.querySelectorAll('ol > li')).filter(li => li.parentElement.parentElement === rem)
      .forEach(li => {
        const c = li.cloneNode(true);
        c.querySelectorAll('ul').forEach(u => u.remove());
        out.notes.push({text: clean(c.textContent),
                        subs: Array.from(li.querySelectorAll('ul > li')).map(x => clean(x.textContent))});
      });
  }
  const c = box.cloneNode(true);
  c.querySelectorAll('table, .remarks-section, h4').forEach(x => x.remove());
  Array.from(c.querySelectorAll('p')).forEach(x => {
    if (/^\s*(以下為|日期)/.test(x.textContent)) x.remove();
  });
  out.extra = clean(c.textContent);
  return out;
}"""


def norm(s):
    return re.sub(r"\s+", "", str(s if s is not None else ""))


def clean(s):
    return re.sub(r"\s+", " ", str(s if s is not None else "")).strip()


def fail(msg):
    print("✗ %s" % msg)
    print("=== 中止：沒有產生或上傳 Excel ===")
    sys.exit(1)


def cur_key(code):
    return (CUR_ORDER.index(code) if code in CUR_ORDER else 99, code)


def parse(d):
    groups = []
    for g in d["data"]:
        rows = []
        for it in g.get("items") or []:
            cur = it.get("currency") or {}
            for code in sorted(cur.keys(), key=cur_key):
                vals = cur[code]
                if vals is None:
                    continue
                if not isinstance(vals, list):
                    fail("資料格式不明：%s｜%s" % (it.get("series"), code))
                raw = [str(v) for v in vals]
                if len(raw) == N_YEARS:
                    kind = "values"
                elif len(raw) == 1 and not PCT.match(raw[0]) and norm(raw[0]) != NA:
                    kind = "message"
                else:
                    fail("數值個數不是 %d 個：%s｜%s｜%d 個" % (N_YEARS, it.get("series"), code, len(raw)))
                rows.append({
                    "series": it.get("series"), "product": clean(it.get("series")),
                    "type": re.sub(r"\s*<br\s*/?>\s*", "", it.get("type") or "").strip(),
                    "cur": code, "raw": raw, "kind": kind})
        groups.append({"name": clean(g.get("name")), "rows": rows})
    return groups


def expected(groups, series, code):
    out = []
    for g in groups:
        for r in g["rows"]:
            if r["series"] == series and r["cur"] == code:
                out.append((norm(g["name"]), [norm(v) for v in r["raw"]]))
    return out


def shown_desktop(st, series, code):
    out, labels = [], None
    for t in st["desktop"]:
        if len(t) < 3 or len(t[0]) < 4:
            return None, None
        labels = [norm(x) for x in t[1]]
        for r in t[2:]:
            if len(r) < 4 or norm(r[0]) != norm(series) or norm(r[2]) != CUR_NAME.get(code, code):
                return None, None
            out.append((norm(t[0][3]), [norm(x) for x in r[3:]]))
    return out, labels


def shown_mobile(st):
    out = []
    for t in st["mobile"]:
        if len(t) < 5 or not t[3]:
            return None
        out.append((norm(t[3][0]), [norm(r[1]) for r in t[4:] if len(r) >= 2]))
    return out


def pick(page, series, code):
    a = b = "?"
    for _ in range(3):
        a = page.evaluate(JS_SET, [0, series])
        page.wait_for_timeout(300)
        b = page.evaluate(JS_SET, [1, code])
        page.wait_for_timeout(300)
        if a == "ok" and b == "ok":
            try:
                page.locator(SEARCH).click(timeout=5000)
            except Exception:
                page.evaluate(JS_CLICK)
            page.wait_for_timeout(500)
            st = page.evaluate(JS_RESULT)
            if st and st["series"] == series and st["confirm"] == code and st["desktop"]:
                return st, ""
        page.wait_for_timeout(1500)
    return None, "選產品 %s｜選貨幣 %s" % (a, b)


def fetch():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False,
                                    args=["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(locale="zh-HK", timezone_id="Asia/Hong_Kong",
                                  viewport={"width": 1366, "height": 900})
        page = ctx.new_page()
        resp = page.goto(URL, wait_until="domcontentloaded", timeout=60000)
        try:
            page.wait_for_load_state("networkidle", timeout=20000)
        except Exception:
            pass
        page.wait_for_timeout(3000)
        print("=== 1. 讀取官方頁 ===")
        print("   HTTP %s｜%s" % (resp.status if resp else "?", page.url))
        d = page.evaluate(JS_DATA)
        if not d or not d.get("data"):
            fail("讀不到頁面資料，官網結構可能已改")
        print("   資料檔：%s" % d["script"])
        groups = parse(d)
        in_data = set(r["series"] for g in groups for r in g["rows"])
        if in_data != set(d["list"]):
            fail("產品清單與資料不對應")

        print("\n=== 2. 逐一核對官網顯示 ===")
        combos = []
        for series in d["list"]:
            codes = sorted(set(r["cur"] for g in groups for r in g["rows"] if r["series"] == series),
                           key=cur_key)
            combos += [(series, c) for c in codes]
        info = {"script": d["script"], "list": d["list"], "labels": None,
                "intro": "", "notes": [], "extras": {}}
        bad = 0
        for series, code in combos:
            st, err = pick(page, series, code)
            if not st:
                bad += 1
                print("   ✗ %s｜%s｜無法顯示結果（%s）" % (clean(series), code, err))
                continue
            exp = expected(groups, series, code)
            desk, labels = shown_desktop(st, series, code)
            mob = shown_mobile(st)
            if desk != exp or mob != exp:
                bad += 1
                print("   ✗ %s｜%s｜不一致" % (clean(series), code))
                print("      資料：%s" % exp)
                print("      桌面：%s" % desk)
                print("      手機：%s" % mob)
                continue
            if info["labels"] is None:
                info["labels"], info["intro"], info["notes"] = labels, st["intro"], st["notes"]
            if st["extra"]:
                info["extras"].setdefault(st["extra"], [])
                if clean(series) not in info["extras"][st["extra"]]:
                    info["extras"][st["extra"]].append(clean(series))
        browser.close()
    print("   核對 %d 組「產品＋貨幣」，%d 組不同（桌面版及手機版）" % (len(combos), bad))
    if bad:
        fail("官網顯示與讀出的資料不一致")
    if not info["labels"] or len(info["labels"]) != N_YEARS:
        fail("官網表頭不是 %d 欄：%s" % (N_YEARS, info["labels"]))
    info["combos"] = len(combos)
    return groups, info


# ---------- 整理資料 ----------
def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def bonus_type(name):
    s = re.sub(r"^\d{4}\s*報告年度之?", "", norm(name))
    s = re.sub(r"之?分紅實現率$", "", s)
    return s or norm(name)


def year_heads(labels):
    out = []
    for i, label in enumerate(labels):
        digits = re.findall(r"\d+", label)
        if i < N_YEARS - 1 and digits[:1] == [str(i + 1)]:
            out.append("保单年度%d" % (i + 1))
        elif i == N_YEARS - 1 and "10" in digits and ("以上" in label or "+" in label):
            out.append("保单年度10+")
        else:
            out.append(label)
    return out


def build_rows(groups, series_list):
    rows, messages, other = [], {}, 0
    for series in series_list:
        for g in groups:
            t = bonus_type(g["name"])
            for r in g["rows"]:
                if r["series"] != series:
                    continue
                fixed = [r["product"], r["type"], CUR_NAME.get(r["cur"], r["cur"]), t]
                if r["kind"] == "message":
                    messages[len(rows)] = clean(r["raw"][0])
                    rows.append(fixed + ["/"] * N_YEARS)
                    continue
                vals = []
                for v in r["raw"]:
                    m = PCT.match(v)
                    if m:
                        vals.append(round(float(m.group(1)) / 100, 4))
                    else:
                        if norm(v) != NA:
                            other += 1
                        vals.append("/")
                rows.append(fixed + vals)
    return rows, messages, other


def mean_or_slash(nums):
    return sum(nums) / len(nums) if nums else "/"


def numbers(rows, n_fixed):
    return [v for r in rows for v in r[n_fixed:] if is_num(v)]


def build_summary(n_fixed, heads, groups):
    n = len(heads)
    width = n_fixed + n
    head = []
    for i in range(n):
        if i == 0:
            head.append("所有年份平均值")
        elif i == n - 1:
            head.append(f"持滿{n - 1}年+平均值")
        elif i == n - 2:
            head.append(f"第{n - 1}年平均值")
        else:
            head.append(f"持滿{i + 1}年平均值")
    cum = []
    for label, sub in groups:
        if not sub:
            continue
        line = [label]
        for start in range(n_fixed, width):
            line.append(mean_or_slash([r[c] for r in sub for c in range(start, width) if is_num(r[c])]))
        cum.append(line)
    all_label, everything = groups[0]
    year_line = [all_label] + [mean_or_slash([r[c] for r in everything if is_num(r[c])])
                               for c in range(n_fixed, width)]
    return {"head": head, "cum": cum, "year_head": list(heads), "year": [year_line]}


def fmt_pct(v):
    return f"{v * 100:g}%" if is_num(v) else "/"


# ---------- 生成 Excel（保誠格式） ----------
def write_row(ws, r, col, values):
    for i, v in enumerate(values):
        ws.cell(row=r, column=col + i, value=v)


def dark(cell):
    cell.fill = DARK
    cell.font = WHITE_BOLD


def light(cell):
    cell.fill = LIGHT
    cell.font = BLACK_BOLD


def heat(cell, v, base):
    if not is_num(v):
        cell.fill = base
        cell.font = GREY_FONT
    elif v >= HIGH_LIMIT:
        cell.fill = GOOD_FILL
        cell.font = GOOD_FONT
    elif v < LOW_LIMIT:
        cell.fill = LOW_FILL
        cell.font = LOW_FONT
    else:
        cell.fill = base


def cat_font(cat):
    s = str(cat)
    if "年金" in s:
        return Font(color="FFC55A11")
    if "儲蓄" in s:
        return Font(color="FF2E7D32")
    return Font(color="FF2F5597")


def units(s):
    return sum(2 if ord(ch) > 255 else 1 for ch in str(s))


def fit_width(values, low, cap):
    longest = max([units(v) for v in values] or [0])
    return min(max(math.ceil(longest * 1.1) + 4, low), cap)


def pct_format(v):
    return "0%" if abs(v * 100 - round(v * 100)) < 1e-9 else "0.0%"


def add_sheet(wb, title, heading, notes, header, rows, messages, summary, head_label, n_fixed, tab):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    last_col = len(header)
    lows = [30, 12, 10, 14]
    caps = [52, 24, 12, 44]
    widths = [fit_width([header[k]] + [row[k] for row in rows], lows[k], caps[k])
              for k in range(n_fixed)]

    ws.cell(row=1, column=1, value=heading)
    ws.cell(row=2, column=1, value=notes)
    for rr in (1, 2):
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
        ws.cell(row=rr, column=1).alignment = LEFT
    ws.cell(row=1, column=1).font = TITLE_FONT
    ws.cell(row=1, column=1).fill = DARK
    ws.cell(row=2, column=1).font = NOTE_FONT
    ws.cell(row=2, column=1).fill = PALE
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 66

    write_row(ws, HEAD_ROW, 1, header)
    for c in range(1, last_col + 1):
        cell = ws.cell(row=HEAD_ROW, column=c)
        dark(cell)
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[HEAD_ROW].height = 42

    rr = FIRST
    for i, row in enumerate(rows):
        first = i == 0 or rows[i - 1][0] != row[0]
        last = i == len(rows) - 1 or rows[i + 1][0] != row[0]
        if first and i > 0:
            ws.cell(row=rr, column=1).fill = SEP_FILL
            ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
            ws.row_dimensions[rr].height = SEP_H
            rr += 1
        border = Border(left=LINE, right=LINE,
                        top=GROUP if first else LINE,
                        bottom=GROUP if last else LINE)
        msg = messages.get(i)
        for c, v in enumerate(row[:n_fixed] if msg else row, 1):
            cell = ws.cell(row=rr, column=c, value=v)
            cell.border = border
            cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
            cell.fill = WHITE
            if c == 1:
                cell.font = NAME_FONT
            elif c == 2:
                cell.font = cat_font(v)
            elif c == 3:
                cell.font = CUR_FONT
            elif c > n_fixed:
                if is_num(v):
                    cell.number_format = pct_format(v)
                heat(cell, v, WHITE)
        if msg:
            cell = ws.cell(row=rr, column=n_fixed + 1, value=msg)
            cell.border = border
            cell.fill = WHITE
            cell.font = MSG_FONT
            cell.alignment = LEFT_FIT
            ws.merge_cells(start_row=rr, start_column=n_fixed + 1, end_row=rr, end_column=last_col)
        ws.row_dimensions[rr].height = ROW_H
        rr += 1
    data_end = rr - 1

    if rows:
        ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(last_col)}{data_end}"

    lc = n_fixed
    r = data_end + 2
    ws.cell(row=r, column=lc, value="平均值摘要").font = SECTION_FONT
    ws.row_dimensions[r].height = ROW_H
    r += 1
    table_start = r
    write_row(ws, r, lc, [head_label] + summary["head"])
    r += 1
    value_rows = []
    for line in summary["cum"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    year_head_row = r
    write_row(ws, r, lc, ["當年平均值"] + summary["year_head"])
    r += 1
    for line in summary["year"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    table_end = r - 1
    for row_i in range(table_start, table_end + 1):
        for c in range(lc, last_col + 1):
            cell = ws.cell(row=row_i, column=c)
            cell.alignment = CENTER
            cell.border = BORDER
    for c in range(lc, last_col + 1):
        dark(ws.cell(row=table_start, column=c))
    dark(ws.cell(row=year_head_row, column=lc))
    for c in range(lc + 1, last_col + 1):
        light(ws.cell(row=year_head_row, column=c))
    for row_i in (table_start, year_head_row):
        ws.row_dimensions[row_i].height = 32
    for vr, line in value_rows:
        label_cell = ws.cell(row=vr, column=lc)
        light(label_cell)
        label_cell.alignment = CENTER_FIT
        ws.row_dimensions[vr].height = ROW_H
        for k, v in enumerate(line[1:]):
            cell = ws.cell(row=vr, column=lc + 1 + k)
            cell.alignment = CENTER_FIT
            cell.number_format = "0.00%"
            heat(cell, v, WHITE)
    note_row = table_end + 1
    ws.cell(row=note_row, column=lc,
            value="「持滿N年平均值」＝保單年度N及之後所有數字的平均；"
                  "「當年平均值」＝該年度一欄所有數字的平均。平均值計入所有貨幣，為固定數字，不會隨篩選變動。")
    ws.merge_cells(start_row=note_row, start_column=lc, end_row=note_row, end_column=last_col)
    ws.cell(row=note_row, column=lc).font = NOTE_FONT
    ws.cell(row=note_row, column=lc).alignment = LEFT
    ws.row_dimensions[note_row].height = 32

    for i in range(1, last_col + 1):
        w = widths[i - 1] if i <= n_fixed else YEAR_W
        if i == lc:
            w = max(w, 18)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=FIRST, column=2)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"


def add_overview(wb, year, date, today, title, rows, n_fixed, info):
    ws = wb.create_sheet("概覽", 0)
    ws.sheet_properties.tabColor = "FFC000"
    ws.sheet_view.showGridLines = False
    last_col = 6
    for i, w in enumerate([36, 18, 18, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def banner(r, value, font, fill, height):
        cell = ws.cell(row=r, column=1, value=value)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        cell.font = font
        cell.fill = fill
        cell.alignment = LEFT
        ws.row_dimensions[r].height = height

    def section(r, value):
        ws.cell(row=r, column=1, value=value).font = SECTION_FONT
        ws.row_dimensions[r].height = ROW_H

    def table_head(r, values):
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            dark(cell)
            cell.alignment = CENTER_FIT
            cell.border = BORDER
        ws.row_dimensions[r].height = ROW_H

    def para(r, value):
        cell = ws.cell(row=r, column=1, value=value)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        cell.font = NOTE_FONT
        cell.alignment = LEFT
        ws.row_dimensions[r].height = max(ROW_H, 17 * math.ceil(units(value) / 110))

    def breakdown(r, label, idx, coloured):
        section(r, label[0])
        r += 1
        table_head(r, [label[1], "整體平均值", "行數", "有數字的格數", "低於 80% 的比例"])
        keys = []
        for row in rows:
            if row[idx] not in keys:
                keys.append(row[idx])
        for key in keys:
            sub = [row for row in rows if row[idx] == key]
            nums = numbers(sub, n_fixed)
            low = sum(1 for v in nums if v < LOW_LIMIT) / len(nums) if nums else "/"
            r += 1
            cell = ws.cell(row=r, column=1, value=key or "（未分類）")
            cell.border = BORDER
            cell.alignment = LEFT_FIT
            cell.font = Font(bold=True, color=cat_font(key).color) if coloured else BLACK_BOLD
            for k, v in enumerate([mean_or_slash(nums), len(sub), len(nums), low]):
                c2 = ws.cell(row=r, column=k + 2, value=v)
                c2.border = BORDER
                c2.alignment = CENTER_FIT
            ws.cell(row=r, column=2).number_format = "0.00%"
            heat(ws.cell(row=r, column=2), mean_or_slash(nums), WHITE)
            ws.cell(row=r, column=5).number_format = "0.0%"
            ws.row_dimensions[r].height = ROW_H
        return r

    banner(1, f"安達人壽分紅實現率概覽（報告年度 {year}）", TITLE_FONT, DARK, 40)
    banner(2, f"資料來源：安達人壽香港網站｜官網標示日期：{date or '未能讀取'}｜更新日期：{today}",
           NOTE_FONT, PALE, ROW_H)

    r = 4
    section(r, "資料概況")
    r += 1
    table_head(r, ["工作表", "產品／版本數", "行數", "有數字的格數", "整體平均值", "低於 80% 的比例"])
    nums = numbers(rows, n_fixed)
    low = sum(1 for v in nums if v < LOW_LIMIT) / len(nums) if nums else "/"
    vals = [title, len({row[0] for row in rows}), len(rows), len(nums), mean_or_slash(nums), low]
    r += 1
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.border = BORDER
        cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
        cell.fill = WHITE
    link = ws.cell(row=r, column=1)
    link.font = LINK_FONT
    link.hyperlink = Hyperlink(ref=link.coordinate, location=f"'{title}'!A1")
    ws.cell(row=r, column=5).number_format = "0.00%"
    heat(ws.cell(row=r, column=5), vals[4], WHITE)
    ws.cell(row=r, column=6).number_format = "0.0%"
    ws.row_dimensions[r].height = ROW_H

    r = breakdown(r + 2, ("按產品種類的整體平均值", "產品種類"), 1, True)
    r = breakdown(r + 2, ("按紅利類型的整體平均值", "紅利類型"), 3, False)

    r += 2
    section(r, "顏色說明")
    samples = [(1.0, "100% 或以上：與銷售時的說明一致或更好"),
               (0.79, "低於 80%：明顯低於銷售時的說明"),
               ("/", "沒有數字：官網顯示「不適用」，或產品於報告年度才推出")]
    for v, desc in samples:
        r += 1
        cell = ws.cell(row=r, column=1, value=v)
        cell.border = BORDER
        cell.alignment = CENTER_FIT
        if is_num(v):
            cell.number_format = "0%"
        heat(cell, v, WHITE)
        d = ws.cell(row=r, column=2, value=desc)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
        d.alignment = LEFT_FIT
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "使用提示")
    tips = ["點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
            "工作表第 3 行有篩選箭嘴，可按產品種類、貨幣或紅利類型篩選。",
            "向右或向下捲動時，標題列和產品名稱會固定不動。",
            "平均值摘要在工作表最底部。",
            "手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。"]
    for t in tips:
        r += 1
        cell = ws.cell(row=r, column=1, value="・" + t)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        cell.font = NOTE_FONT
        cell.alignment = LEFT_FIT
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "官網說明及備註（原文）")
    if info["body"]:
        r += 1
        para(r, info["body"])
    for i, n in enumerate(info["notes"], 1):
        r += 1
        para(r, "%d. %s" % (i, n["text"]))
        for sub in n["subs"]:
            r += 1
            para(r, "　　• %s" % sub)
    for text, prods in info["extras"].items():
        r += 1
        para(r, "%s：%s" % ("、".join(prods), text))


# ---------- Notion（與保誠相同） ----------
def check(r, what):
    if r.status_code >= 300:
        print(f"{what} 失敗：{r.status_code} {r.text[:400]}")
        raise SystemExit(1)
    return r.json()


def resolve_source(raw):
    m = re.search(r"[0-9a-fA-F]{32}", raw.split("?")[0].replace("-", ""))
    ident = m.group(0) if m else raw

    def first_source(db_id):
        r = requests.get(f"{API}/databases/{db_id}", headers=HEADERS, timeout=30)
        if r.status_code == 200:
            ds = r.json().get("data_sources") or []
            return ds[0]["id"] if ds else None
        return None

    found = first_source(ident)
    if found:
        return found
    r = requests.get(f"{API}/data_sources/{ident}", headers=HEADERS, timeout=30)
    if r.status_code == 200:
        return ident
    status = r.status_code
    r = requests.get(f"{API}/blocks/{ident}/children?page_size=100", headers=HEADERS, timeout=30)
    if r.status_code == 200:
        for b in r.json().get("results", []):
            if b.get("type") == "child_database":
                found = first_source(b["id"])
                if found:
                    return found
    print(f"找不到安達人壽的年份資料庫（狀態 {status}）。401＝金鑰不正確；"
          "404＝頁面未連接「AIA 自動更新」或 NOTION_DATABASE_ID 不正確。")
    raise SystemExit(1)


def query_pages(source):
    r = requests.post(f"{API}/data_sources/{source}/query", headers=HEADERS, json={"page_size": 100}, timeout=30)
    return check(r, "讀取年份資料庫").get("results", [])


def page_title(page):
    for prop in page.get("properties", {}).values():
        if prop.get("type") == "title":
            return "".join(t.get("plain_text", "") for t in prop.get("title", []))
    return ""


def file_blocks(page_id):
    r = requests.get(f"{API}/blocks/{page_id}/children?page_size=100", headers=HEADERS, timeout=30)
    out = []
    for block in check(r, "讀取頁面內容").get("results", []):
        if block["type"] == "file":
            info = block["file"]
            url = (info.get("file") or info.get("external") or {}).get("url")
            out.append((block["id"], info.get("name", ""), url))
    return out


def upload_file(page_id, filename, content):
    old_auto = [bid for bid, name, _ in file_blocks(page_id) if AUTO in name]
    r = requests.post(f"{API}/file_uploads", headers=HEADERS,
                      json={"mode": "single_part", "filename": filename, "content_type": XLSX}, timeout=30)
    upload_id = check(r, "建立上傳")["id"]
    r = requests.post(f"{API}/file_uploads/{upload_id}/send",
                      headers={"Authorization": f"Bearer {TOKEN}", "Notion-Version": VERSION},
                      files={"file": (filename, content, XLSX)}, timeout=120)
    check(r, "傳送檔案")
    block = {"object": "block", "type": "file",
             "file": {"type": "file_upload", "file_upload": {"id": upload_id}}}
    r = requests.patch(f"{API}/blocks/{page_id}/children", headers=HEADERS, json={"children": [block]}, timeout=60)
    check(r, "加入檔案到頁面")
    for bid in old_auto:
        r = requests.delete(f"{API}/blocks/{bid}", headers=HEADERS, timeout=30)
        check(r, "刪除上一次的自動生成檔案")
    print(f"已上傳：{filename}（刪除舊的自動生成檔案 {len(old_auto)} 個）")


# ---------- 主流程 ----------
def main():
    if UPLOAD and (not TOKEN or not TARGET):
        fail("未設定 NOTION_TOKEN 或 NOTION_DATABASE_ID")

    groups, info = fetch()
    years = set(re.findall(r"(\d{4})\s*報告年度", " ".join(g["name"] for g in groups)))
    if len(years) != 1:
        fail("報告年度不一致：%s" % sorted(years))
    year = years.pop()
    m = re.search(r"日期[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)", info["intro"])
    date = m.group(1) if m else ""
    info["body"] = clean(info["intro"][m.end():]) if m else info["intro"]
    today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d")

    rows, messages, other = build_rows(groups, info["list"])
    heads = year_heads(info["labels"])
    nums = numbers(rows, N_FIXED)

    print("\n=== 3. 讀取結果 ===")
    print("   報告年度 %s｜官網日期 %s｜產品 %d 個" % (year, date or "未能讀取", len(info["list"])))
    print("   官網表頭：%s" % "｜".join(info["labels"]))
    print("   Excel 表頭：%s" % "｜".join(heads))
    types = []
    for row in rows:
        if row[3] not in types:
            types.append(row[3])
    for t in types:
        sub = [(i, row) for i, row in enumerate(rows) if row[3] == t]
        n = numbers([row for _, row in sub], N_FIXED)
        print("   %s：產品 %d 個｜%d 行｜數值 %d 個｜平均 %s｜最高 %s｜最低 %s｜低於 80%% 的 %d 個｜未能提供 %d 行" % (
            t, len({row[0] for _, row in sub}), len(sub), len(n),
            "%.2f%%" % (sum(n) / len(n) * 100) if n else "/",
            fmt_pct(max(n)) if n else "/", fmt_pct(min(n)) if n else "/",
            len([x for x in n if x < LOW_LIMIT]), len([i for i, _ in sub if i in messages])))
    if not nums:
        fail("沒有讀到任何數值")
    print("   合計：%d 行｜數值 %d 個｜平均 %.2f%%｜最高 %s｜最低 %s｜低於 80%% 的 %d 個｜高於 100%% 的 %d 個｜未能提供 %d 行" % (
        len(rows), len(nums), sum(nums) / len(nums) * 100, fmt_pct(max(nums)), fmt_pct(min(nums)),
        len([x for x in nums if x < LOW_LIMIT]), len([x for x in nums if x > HIGH_LIMIT]), len(messages)))
    if other:
        fail("有 %d 格既不是百分比也不是「不適用」，官網格式可能已改" % other)
    if len(rows) < 40 or len(nums) < 80:
        fail("讀到的資料太少，官網可能改了版面")
    print("   低於 80% 的行：")
    for row in rows:
        if any(is_num(v) and v < LOW_LIMIT for v in row[N_FIXED:]):
            print("      %s｜%s｜%s｜%s" % (row[0], row[2], row[3], " ".join(fmt_pct(v) for v in row[N_FIXED:])))
    for i, text in messages.items():
        print("   未能提供｜%s｜%s｜%s｜%s" % (rows[i][0], rows[i][2], rows[i][3], text))

    print("\n=== 4. 官網說明及備註 ===")
    print("   說明：%s" % info["body"])
    for i, n in enumerate(info["notes"], 1):
        print("   %d. %s" % (i, n["text"]))
        for sub in n["subs"]:
            print("      • %s" % sub)
    for text, prods in info["extras"].items():
        print("   額外備註｜%s：%s" % ("、".join(prods), text))

    f_groups = [("所有分紅", rows)]
    for t in types:
        sub = [row for row in rows if row[3] == t]
        if numbers(sub, N_FIXED):
            f_groups.append((t, sub))
    summary = build_summary(N_FIXED, heads, f_groups)
    notes = ("分紅實現率的定義及計算基準以安達人壽網站的說明為準（原文見「概覽」）。"
             "100% 即與銷售時的說明一致，高於 100% 即比說明多。\n"
             "綠色＝100% 或以上；紅色＝低於 80%；「/」＝官網顯示「不適用」。\n"
             f"資料來源：安達人壽香港網站（報告年度 {year}，官網標示日期 {date or '未能讀取'}）｜更新日期：{today}")
    title = f"{year}分紅實現率"
    wb = Workbook()
    wb.remove(wb.active)
    add_sheet(wb, title, f"安達人壽分紅實現率（報告年度 {year}）", notes,
              ["产品名称", "产品种类", "货币", "红利类型"] + heads, rows, messages, summary,
              "分紅類型", N_FIXED, "2E75B6")
    add_overview(wb, year, date, today, title, rows, N_FIXED, info)
    buf = io.BytesIO()
    wb.save(buf)
    content = buf.getvalue()
    local = "chubb_fulfillment_%s.xlsx" % year
    with open(local, "wb") as fh:
        fh.write(content)
    print("\n=== 5. Excel ===")
    print("   %s｜%d KB｜工作表：%s" % (local, len(content) // 1024, "、".join(wb.sheetnames)))

    print("\n=== 6. 上傳到 Notion ===")
    if not UPLOAD:
        print("   試跑：沒有上傳")
        print("完成")
        return
    source = resolve_source(TARGET)
    pages = query_pages(source)
    page = next((p for p in pages if str(year) in page_title(p)), None)
    if page is None:
        body = {"parent": {"type": "data_source_id", "data_source_id": source},
                "properties": {"title": {"title": [{"text": {"content": str(year)}}]}}}
        r = requests.post(f"{API}/pages", headers=HEADERS, json=body, timeout=30)
        page = check(r, "建立年度頁面")
        print(f"已建立頁面：{year}")
    else:
        print("使用現有頁面：", page_title(page))
    upload_file(page["id"], f"安達人壽分紅實現率（{year}年）{AUTO}.xlsx", content)
    print("完成")


if __name__ == "__main__":
    main()

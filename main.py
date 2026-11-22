import re
import sys
import math
import datetime
from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

URL = "https://www.chubblife.com.hk/zh/customer-service/fulfillment-ratios-of-dividend.html"
SEARCH = "#fulfillment-section .search-btn.search"
CUR_NAME = {"HKD": "港元", "USD": "美元"}
CUR_ORDER = ["HKD", "USD"]
NA = "不適用"
PCT = re.compile(r"^\s*(\d+(?:\.\d+)?)\s*%\s*$")

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
    print("=== 中止：沒有產生 Excel ===")
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
                if len(raw) == 11:
                    kind = "values"
                elif len(raw) == 1 and not PCT.match(raw[0]) and norm(raw[0]) != NA:
                    kind = "message"
                else:
                    fail("數值個數不是 11 個：%s｜%s｜%d 個" % (it.get("series"), code, len(raw)))
                rows.append({
                    "series": it.get("series"), "product": clean(it.get("series")),
                    "eng": clean(it.get("eng")),
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
        info = {"script": d["script"], "labels": None, "intro": "", "notes": [], "extras": {}}
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
    if not info["labels"] or len(info["labels"]) != 11:
        fail("官網表頭不是 11 欄：%s" % info["labels"])
    info["combos"] = len(combos)
    info["products"] = len(d["list"])
    return groups, info


def stats(g):
    nums, na, other, msg, low_rows = [], 0, 0, 0, []
    for r in g["rows"]:
        if r["kind"] == "message":
            msg += 1
            continue
        mine = []
        for v in r["raw"]:
            m = PCT.match(v)
            if m:
                mine.append(float(m.group(1)))
            elif norm(v) == NA:
                na += 1
            else:
                other += 1
        nums += mine
        if mine and min(mine) < 80:
            low_rows.append(r)
    return {"products": len(set(r["series"] for r in g["rows"])), "rows": len(g["rows"]),
            "n": len(nums), "avg": sum(nums) / len(nums) if nums else 0,
            "max": max(nums) if nums else 0, "min": min(nums) if nums else 0,
            "low": len([x for x in nums if x < 80]), "high": len([x for x in nums if x >= 100]),
            "over": len([x for x in nums if x > 100]),
            "na": na, "other": other, "msg": msg, "low_rows": low_rows, "nums": nums}


THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
GREEN = PatternFill("solid", fgColor="C6EFCE")
RED = PatternFill("solid", fgColor="FFC7CE")
HEAD = PatternFill("solid", fgColor="1F3864")
SEP = PatternFill("solid", fgColor="F2F2F2")
BOLD = Font(bold=True)
WHITE = Font(bold=True, color="FFFFFF")
PURPLE = Font(bold=True, color="7030A0")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def sheet_name(name, i, used):
    s = re.sub(r"^\d{4}\s*報告年度之", "", name)
    s = re.sub(r"之?分紅實現率$", "", s)
    s = re.sub(r"[\\/*?:\[\]]", "", s).strip()[:31] or "第%d組" % i
    if s in used:
        s = ("%s%d" % (s[:29], i))
    used.add(s)
    return s


def write_group(ws, g, labels):
    ncol = 4 + len(labels)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    c = ws.cell(1, 1, g["name"])
    c.font, c.alignment = Font(bold=True, size=13), LEFT
    ws.row_dimensions[1].height = 24
    for i, h in enumerate(["產品名稱", "英文名稱", "產品類別", "貨幣"] + labels, 1):
        c = ws.cell(2, i, h)
        c.font, c.fill, c.alignment, c.border = WHITE, HEAD, CENTER, BORDER
    ws.row_dimensions[2].height = 34
    r, prev = 3, None
    for row in g["rows"]:
        if prev is not None and row["series"] != prev:
            for i in range(1, ncol + 1):
                ws.cell(r, i).fill = SEP
            ws.row_dimensions[r].height = 6
            r += 1
        prev = row["series"]
        for i, (v, font, al) in enumerate([
                (row["product"], BOLD, LEFT), (row["eng"], None, LEFT), (row["type"], None, CENTER),
                (CUR_NAME.get(row["cur"], row["cur"]), PURPLE, CENTER)], 1):
            c = ws.cell(r, i, v)
            c.alignment, c.border = al, BORDER
            if font:
                c.font = font
        if row["kind"] == "message":
            ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=ncol)
            c = ws.cell(r, 5, row["raw"][0])
            c.font, c.alignment = Font(italic=True, color="595959"), LEFT
            ws.row_dimensions[r].height = 30
        else:
            for i, v in enumerate(row["raw"]):
                c = ws.cell(r, 5 + i)
                c.alignment, c.border = CENTER, BORDER
                m = PCT.match(v)
                if m:
                    x = float(m.group(1))
                    c.value = x / 100
                    c.number_format = "0%" if x == int(x) else "0.0#%"
                    if x >= 100:
                        c.fill = GREEN
                    elif x < 80:
                        c.fill = RED
                elif norm(v) == NA:
                    c.value = "/"
                else:
                    c.value = v
        r += 1
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 8
    for i in range(5, ncol + 1):
        ws.column_dimensions[get_column_letter(i)].width = 11
    ws.freeze_panes = "E3"


def para(ws, r, text, font=None):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    c = ws.cell(r, 1, text)
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    if font:
        c.font = font
    ws.row_dimensions[r].height = max(16, 16 * math.ceil(len(text) / 60.0))
    return r + 1


def write_overview(ws, groups, names, info, year, date, now):
    ws.title = "概覽"
    ws.column_dimensions["A"].width = 24
    for i in range(2, 11):
        ws.column_dimensions[get_column_letter(i)].width = 13
    r = para(ws, 1, "安達人壽（香港）分紅實現率｜%s 報告年度" % year, Font(bold=True, size=14))
    r += 1
    for k, v in [("官網標示日期", date or "未能讀取"), ("抓取時間（香港）", now),
                 ("來源頁", URL), ("資料檔", info["script"]),
                 ("產品數目", "%d 個" % info["products"]),
                 ("核對", "%d 組「產品＋貨幣」與官網桌面版及手機版表格一致" % info["combos"]),
                 ("圖例", "綠底：100% 或以上｜紅底：低於 80%｜/：官網顯示「不適用」")]:
        ws.cell(r, 1, k).font = BOLD
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
        ws.cell(r, 2, v).alignment = LEFT
        r += 1
    r += 1
    for i, h in enumerate(["工作表", "產品數", "行數", "數值個數", "平均", "最高", "最低",
                           "低於 80%", "100% 或以上", "未能提供"], 1):
        c = ws.cell(r, i, h)
        c.font, c.fill, c.alignment, c.border = WHITE, HEAD, CENTER, BORDER
    r += 1
    for g, name in zip(groups, names):
        s = stats(g)
        vals = [name, s["products"], s["rows"], s["n"], "%.2f%%" % s["avg"], "%g%%" % s["max"],
                "%g%%" % s["min"], s["low"], s["high"], s["msg"]]
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v)
            c.alignment, c.border = (LEFT if i == 1 else CENTER), BORDER
        r += 1
    r += 1
    r = para(ws, r, "官網說明", BOLD)
    r = para(ws, r, info["body"])
    r += 1
    r = para(ws, r, "官網備註", BOLD)
    for i, n in enumerate(info["notes"], 1):
        r = para(ws, r, "%d. %s" % (i, n["text"]))
        for sub in n["subs"]:
            r = para(ws, r, "　　• %s" % sub)
    if info["extras"]:
        r += 1
        r = para(ws, r, "個別產品的額外備註", BOLD)
        for text, prods in info["extras"].items():
            r = para(ws, r, "%s：%s" % ("、".join(prods), text))


def main():
    groups, info = fetch()
    years = set(re.findall(r"(\d{4})\s*報告年度", " ".join(g["name"] for g in groups)))
    if len(years) != 1:
        fail("報告年度不一致：%s" % sorted(years))
    year = years.pop()
    m = re.search(r"日期[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)", info["intro"])
    date = m.group(1) if m else ""
    info["body"] = clean(info["intro"][m.end():]) if m else info["intro"]
    now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d %H:%M")

    print("\n=== 3. 資料統計 ===")
    print("   報告年度 %s｜官網日期 %s｜產品 %d 個" % (year, date or "未能讀取", info["products"]))
    print("   表頭：%s" % "｜".join(info["labels"]))
    used, names, total = set(), [], []
    for i, g in enumerate(groups, 1):
        s = stats(g)
        names.append(sheet_name(g["name"], i, used))
        total += s["nums"]
        print("\n   --- %s（工作表「%s」）---" % (g["name"], names[-1]))
        print("   產品 %d 個｜%d 行｜數值 %d 個｜平均 %.2f%%｜最高 %g%%｜最低 %g%%" % (
            s["products"], s["rows"], s["n"], s["avg"], s["max"], s["min"]))
        print("   低於 80%% 的 %d 個｜100%% 或以上 %d 個｜高於 100%% 的 %d 個｜不適用 %d 格｜未能提供 %d 行｜其他文字 %d 格" % (
            s["low"], s["high"], s["over"], s["na"], s["msg"], s["other"]))
        for r in s["low_rows"]:
            print("   低值｜%s｜%s｜%s" % (r["product"], CUR_NAME.get(r["cur"], r["cur"]),
                                       " ".join("/" if norm(v) == NA else v for v in r["raw"])))
        for r in g["rows"]:
            if r["kind"] == "message":
                print("   未能提供｜%s｜%s｜%s" % (r["product"], CUR_NAME.get(r["cur"], r["cur"]), r["raw"][0]))
    print("\n   合計：%d 行｜數值 %d 個｜最高 %g%%｜最低 %g%%｜低於 80%% 的 %d 個｜高於 100%% 的 %d 個" % (
        sum(len(g["rows"]) for g in groups), len(total), max(total), min(total),
        len([x for x in total if x < 80]), len([x for x in total if x > 100])))

    print("\n=== 4. 官網說明及備註 ===")
    print("   說明：%s" % info["body"])
    for i, n in enumerate(info["notes"], 1):
        print("   %d. %s" % (i, n["text"]))
        for sub in n["subs"]:
            print("      • %s" % sub)
    print("   個別產品的額外備註 %d 段" % len(info["extras"]))
    for text, prods in info["extras"].items():
        print("   %s：%s" % ("、".join(prods), text))

    wb = Workbook()
    write_overview(wb.active, groups, names, info, year, date, now)
    for g, name in zip(groups, names):
        write_group(wb.create_sheet(name), g, info["labels"])
    path = "chubb_fulfillment_%s.xlsx" % year
    wb.save(path)
    print("\n=== 5. 已產生 %s（未上傳 Notion）===" % path)
    print("=== 完成 ===")


if __name__ == "__main__":
    main()

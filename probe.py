import re
from collections import Counter
from playwright.sync_api import sync_playwright

URL = "https://www.chubblife.com.hk/zh/customer-service/fulfillment-ratios-of-dividend.html"
CASES = [
    ("「星鑽」儲蓄壽險計劃", "HKD"),
    ("「星鑽」儲蓄壽險計劃", "USD"),
    ("安達開創未來儲蓄計劃II", "USD"),
    ("「智策人生」儲蓄保障計劃", "HKD"),
    ("安達靈「活」延期年金計劃", "USD"),
]
PCT = re.compile(r"^\s*(\d+(?:\.\d+)?)\s*%\s*$")

JS_DATA = r"""() => {
  const el = document.querySelector('#fulfillment-section');
  const vm = el && el.__vue__;
  if (!vm) return null;
  return JSON.parse(JSON.stringify({
    keys: Object.keys(vm.$data),
    list: vm.$data.fulfillmentSeriesList,
    data: vm.$data.data}));
}"""

JS_SCRIPT = r"""async () => {
  const s = Array.from(document.querySelectorAll('script[src]')).map(x => x.src)
    .find(x => /hk-zh\.js/.test(x));
  if (!s) return null;
  try { const r = await fetch(s); return {url: s, text: await r.text()}; }
  catch (e) { return {url: s, text: 'ERR ' + e}; }
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
  if (!b) return '找不到搜尋按鈕';
  b.click();
  return 'ok';
}"""

JS_STATE = r"""() => {
  const clean = s => (s || '').replace(/\s+/g, ' ').trim();
  const el = document.querySelector('#fulfillment-section');
  if (!el) return null;
  const d = el.__vue__ ? el.__vue__.$data : {};
  let result = '';
  try { result = JSON.stringify(d.result) || ''; } catch (e) { result = 'ERR ' + e; }
  const copy = el.cloneNode(true);
  copy.querySelectorAll('select').forEach(s =>
    Array.from(s.options).forEach((o, i) => { if (i > 1) o.remove(); }));
  return {
    series: d.fulfillmentSeries, currency: d.currency, confirm: d.currencyConfirm,
    remarks: d.showRemarks, result: result,
    tables: Array.from(el.querySelectorAll('table')).map(t =>
      Array.from(t.rows).map(r => Array.from(r.cells).map(x =>
        clean(x.textContent) + (x.colSpan > 1 ? '〔跨' + x.colSpan + '欄〕' : '') +
        (x.rowSpan > 1 ? '〔跨' + x.rowSpan + '行〕' : '')))),
    text: clean(copy.textContent).slice(0, 6000),
    html: copy.outerHTML.replace(/\s+/g, ' ').slice(0, 14000)
  };
}"""


def short(e, n=300):
    return re.sub(r"\s+", " ", str(e))[:n]


def load(ctx):
    page = ctx.new_page()
    page.goto(URL, wait_until="domcontentloaded", timeout=60000)
    try:
        page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:
        pass
    page.wait_for_timeout(3000)
    return page


def show(v):
    return "/" if v == "不適用" else str(v)


def dump_data(d):
    print("   Vue 欄位：%s" % "、".join(d["keys"]))
    print("   產品清單 %d 個｜資料分組 %d 組（以下「/」代表官網原文「不適用」）" % (
        len(d["list"]), len(d["data"])))
    lengths, texts, nums, member = Counter(), Counter(), [], {}
    for gi, g in enumerate(d["data"], 1):
        items = g.get("items") or []
        print("\n   --- 第 %d 組：%s｜%d 個產品｜分組欄位：%s ---" % (
            gi, g.get("name"), len(items), "、".join(g.keys())))
        for k, v in g.items():
            if k not in ("name", "items"):
                print("   分組的其他欄位 %s：%s" % (k, short(v, 1500)))
        for it in items:
            name = it.get("series")
            member.setdefault(name, []).append(gi)
            kind = re.sub(r"\s*<br\s*/?>\s*", "", it.get("type") or "")
            extra = [k for k in it.keys() if k not in ("series", "eng", "type", "currency")]
            parts = []
            for cur, vals in (it.get("currency") or {}).items():
                if vals is None:
                    parts.append("%s 沒有" % cur)
                    continue
                if not isinstance(vals, list):
                    parts.append("%s 非陣列：%s" % (cur, short(vals, 200)))
                    continue
                lengths[len(vals)] += 1
                for v in vals:
                    m = PCT.match(str(v))
                    if m:
                        nums.append(float(m.group(1)))
                    else:
                        texts[str(v)[:120]] += 1
                parts.append("%s[%d]：%s" % (cur, len(vals), " ".join(show(v) for v in vals)))
            print("   %s｜%s｜%s%s" % (name, kind, "｜".join(parts),
                                     "｜其他欄位 " + "、".join(extra) if extra else ""))
    print("\n   --- 統計 ---")
    print("   每行數值個數：%s" % "；".join("%d 個 × %d 行" % (k, v) for k, v in sorted(lengths.items())))
    if nums:
        print("   百分比 %d 個｜最高 %.2f%%｜最低 %.2f%%｜低於 80%% 的 %d 個｜高於 100%% 的 %d 個" % (
            len(nums), max(nums), min(nums),
            len([x for x in nums if x < 80]), len([x for x in nums if x > 100])))
    print("   非百分比的內容：")
    for k, v in texts.most_common(30):
        print("      %d 次｜%s" % (v, k))
    combo = Counter(tuple(v) for v in member.values())
    print("   產品所屬分組：%s" % "；".join(
        "第 %s 組 × %d 個產品" % ("、".join(str(x) for x in k), v) for k, v in sorted(combo.items())))
    missing = [x for x in d["list"] if x not in member]
    print("   清單有但資料沒有的產品 %d 個：%s" % (len(missing), "｜".join(missing)))
    extra = [x for x in member if x not in d["list"]]
    print("   資料有但清單沒有的產品 %d 個：%s" % (len(extra), "｜".join(extra)))


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False,
                                    args=["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(locale="zh-HK", timezone_id="Asia/Hong_Kong",
                                  viewport={"width": 1366, "height": 900})
        page = load(ctx)

        print("=== 1. 全部資料 ===")
        d = page.evaluate(JS_DATA)
        if d:
            dump_data(d)
        else:
            print("   讀不到 Vue 資料")

        print("\n=== 2. hk-zh.js 的顯示程式碼 ===")
        s = page.evaluate(JS_SCRIPT)
        if s:
            text = s["text"]
            print("   %s｜%d 字" % (s["url"], len(text)))
            pos = text.find("showRemarks")
            start = max(0, pos - 300) if pos >= 0 else max(0, len(text) - 6000)
            print("   片段（由第 %d 字起）：%s" % (start, short(text[start:start + 6000], 6000)))
        else:
            print("   找不到 hk-zh.js")
        page.close()

        for n, (name, cur) in enumerate(CASES, 1):
            print("\n=== 3.%d %s｜%s ===" % (n, name, cur))
            try:
                page = load(ctx)
                print("   選產品：%s" % page.evaluate(JS_SET, [0, name]))
                page.wait_for_timeout(1500)
                print("   選貨幣：%s" % page.evaluate(JS_SET, [1, cur]))
                page.wait_for_timeout(1500)
                try:
                    page.locator("#fulfillment-section .search-btn.search").click(timeout=5000)
                    print("   按搜尋：ok")
                except Exception as e:
                    print("   按搜尋失敗，改用程式點擊：%s｜%s" % (page.evaluate(JS_CLICK), short(e, 150)))
                page.wait_for_timeout(3000)
                st = page.evaluate(JS_STATE)
                if not st:
                    print("   找不到 fulfillment-section")
                    page.close()
                    continue
                print("   Vue 狀態：產品 %s｜貨幣 %s｜已確認貨幣 %s｜顯示備註 %s" % (
                    st["series"], st["currency"], st["confirm"], st["remarks"]))
                print("   Vue result（%d 字）：%s" % (len(st["result"]), st["result"][:3000]))
                print("   區塊內表格 %d 個" % len(st["tables"]))
                for ti, rows in enumerate(st["tables"], 1):
                    print("   表格 %d（%d 行）：" % (ti, len(rows)))
                    for r in rows[:30]:
                        print("      %s" % "｜".join(r)[:1500])
                print("   區塊文字：%s" % st["text"])
                if n in (1, 3):
                    print("   區塊 HTML：%s" % st["html"])
            except Exception as e:
                print("   操作出錯：%s" % short(e))
            try:
                page.close()
            except Exception:
                pass
        browser.close()
    print("\n=== 探查完成 ===")


if __name__ == "__main__":
    main()

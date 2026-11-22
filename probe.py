import re
from playwright.sync_api import sync_playwright

URL = "https://www.chubblife.com.hk/zh/customer-service/fulfillment-ratios-of-dividend.html"
TRACK = re.compile(r"google|doubleclick|facebook|adsrvr|linkedin|yimg|omtrdc|analytics|demdex", re.I)
SKIP_KIND = ("image", "font", "stylesheet", "media")
NAMES = ["安達開創未來儲蓄計劃", "智策人生", "星鑽"]

JS_SCRIPT_URLS = r"""() => Array.from(document.querySelectorAll('script[src]')).map(s => s.src)
  .filter(x => /chubblife\.com\.hk/.test(x) && !/jquery|granite|vue@|accordion/i.test(x))"""

JS_FETCH = r"""async (u) => {
  try { const r = await fetch(u); return await r.text(); } catch (e) { return 'ERR ' + e; }
}"""

JS_VUE = r"""() => {
  const out = [];
  let n = document.querySelector('#fulfillment-section');
  for (let d = 0; n && d < 8; d++, n = n.parentElement) {
    if (!n.__vue__) continue;
    const vm = n.__vue__;
    const data = vm.$data || {};
    const sizes = Object.keys(data).map(k => {
      let len = -1;
      try { len = (JSON.stringify(data[k]) || '').length; } catch (e) {}
      return k + '（' + (Array.isArray(data[k]) ? '陣列 ' + data[k].length + ' 項' : typeof data[k]) + '，' + len + ' 字）';
    });
    let json = '';
    try { json = JSON.stringify(data); } catch (e) { json = 'ERR ' + e; }
    out.push({where: n.tagName + '#' + n.id, sizes: sizes, size: json.length, head: json.slice(0, 3500),
              methods: Object.keys(vm.$options.methods || {}),
              computed: Object.keys(vm.$options.computed || {})});
  }
  return out;
}"""

JS_PICK = r"""(idx) => {
  const s = document.querySelector('select[name="fulfillment-radios-type"]');
  if (!s || idx >= s.options.length) return null;
  s.selectedIndex = idx;
  s.dispatchEvent(new Event('input', {bubbles: true}));
  s.dispatchEvent(new Event('change', {bubbles: true}));
  return (s.options[idx].textContent || '').replace(/\s+/g, ' ').trim();
}"""

JS_SECTION = r"""() => {
  const clean = s => (s || '').replace(/\s+/g, ' ').trim();
  const sec = document.querySelector('#fulfillment-section');
  if (!sec) return null;
  const c = sec.cloneNode(true);
  c.querySelectorAll('select').forEach(s =>
    Array.from(s.options).forEach((o, i) => { if (i > 2) o.remove(); }));
  return {
    html: c.outerHTML.replace(/\s+/g, ' ').slice(0, 7000),
    text: clean(c.textContent).slice(0, 3000),
    selects: Array.from(sec.querySelectorAll('select')).map(s =>
      (s.name || s.id || '?') + '：' + s.options.length + ' 項｜' +
      Array.from(s.options).slice(0, 6).map(o =>
        clean(o.textContent) + '=' + o.value + (o.disabled ? '(停用)' : '')).join('｜')),
    controls: Array.from(sec.querySelectorAll('input, button, a, label')).slice(0, 30).map(e =>
      e.tagName + '｜type=' + (e.type || '') + '｜name=' + (e.name || '') + '｜value=' +
      (e.value || '') + '｜' + clean(e.textContent).slice(0, 60)),
    tables: Array.from(sec.querySelectorAll('table')).map(t => ({
      rows: t.rows.length,
      cols: Math.max(0, ...Array.from(t.rows).map(r => r.cells.length)),
      text: clean(t.textContent).slice(0, 1800)})),
    all_tables: document.querySelectorAll('table').length,
    links: Array.from(sec.querySelectorAll('a[href]')).slice(0, 20).map(a =>
      clean(a.textContent).slice(0, 60) + '｜' + a.href)
  };
}"""

JS_CHOICES = r"""() => {
  const clean = s => (s || '').replace(/\s+/g, ' ').trim();
  const sec = document.querySelector('#fulfillment-section');
  if (!sec) return {kind: 'none', labels: [], idx: []};
  const sels = Array.from(sec.querySelectorAll('select'));
  if (sels.length >= 2) {
    const s = sels[1];
    const idx = Array.from(s.options).map((o, i) => i)
      .filter(i => !s.options[i].disabled && s.options[i].value !== '');
    return {kind: 'select', idx: idx, labels: idx.map(i => clean(s.options[i].textContent))};
  }
  const radios = Array.from(sec.querySelectorAll('input[type=radio]'));
  if (radios.length) {
    return {kind: 'radio', idx: radios.map((r, i) => i),
            labels: radios.map(r => r.value || clean((r.closest('label') || r.parentElement).textContent))};
  }
  return {kind: 'none', labels: [], idx: []};
}"""

JS_APPLY = r"""(a) => {
  const sec = document.querySelector('#fulfillment-section');
  if (a[0] === 'select') {
    const s = sec.querySelectorAll('select')[1];
    s.selectedIndex = a[1];
    s.dispatchEvent(new Event('input', {bubbles: true}));
    s.dispatchEvent(new Event('change', {bubbles: true}));
    return s.value;
  }
  const r = sec.querySelectorAll('input[type=radio]')[a[1]];
  r.click();
  return r.value;
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


def show_vue(page, head):
    try:
        items = page.evaluate(JS_VUE)
    except Exception as e:
        print("   讀取 Vue 資料失敗：%s" % short(e))
        return
    print("   Vue 元件 %d 個" % len(items))
    for v in items:
        print("   位置 %s｜資料共 %d 字" % (v["where"], v["size"]))
        print("      欄位：%s" % "；".join(v["sizes"]))
        print("      方法：%s" % "、".join(v["methods"]))
        print("      計算欄位：%s" % "、".join(v["computed"]))
        print("      資料開頭：%s" % v["head"][:head])


def show_section(page, with_html):
    try:
        s = page.evaluate(JS_SECTION)
    except Exception as e:
        print("   讀取區塊失敗：%s" % short(e))
        return
    if not s:
        print("   找不到 fulfillment-section")
        return
    print("   全頁表格 %d 個｜區塊內表格 %d 個｜區塊內連結 %d 個" % (
        s["all_tables"], len(s["tables"]), len(s["links"])))
    for x in s["selects"]:
        print("   選單：%s" % x)
    for x in s["controls"]:
        print("   控件：%s" % x)
    for t in s["tables"]:
        print("   表格 %d 行 × %d 欄｜%s" % (t["rows"], t["cols"], t["text"]))
    for x in s["links"]:
        print("   連結：%s" % x[:300])
    print("   區塊文字：%s" % s["text"])
    if with_html:
        print("   區塊 HTML：%s" % s["html"])


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False,
                                    args=["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(locale="zh-HK", timezone_id="Asia/Hong_Kong",
                                  viewport={"width": 1366, "height": 900})
        page = load(ctx)

        print("=== 1. 網站程式檔 ===")
        for u in page.evaluate(JS_SCRIPT_URLS):
            text = page.evaluate(JS_FETCH, u)
            print("   %s｜%d 字" % (u[:200], len(text)))
            for key in ["fulfil"] + NAMES:
                hits = [m.start() for m in re.finditer(re.escape(key), text, re.I)]
                print("      「%s」%d 次" % (key, len(hits)))
                for pos in hits[:3]:
                    print("         …%s…" % short(text[max(0, pos - 250):pos + 650], 900))
            files = sorted(set(re.findall(r"[\w\-./:%]+\.(?:pdf|json|xlsx?|csv)", text, re.I)))
            print("      檔案路徑 %d 個：%s" % (len(files), "｜".join(files[:30])[:1500]))
            print("      百分比數字 %d 個" % len(re.findall(r"\d+(?:\.\d+)?\s*%", text)))

        print("\n=== 2. 選產品前的 Vue 資料 ===")
        show_vue(page, 3500)
        page.close()

        first = True
        for n, idx in enumerate((2, 24, 44), 1):
            print("\n=== 3.%d 選第 [%d] 項 ===" % (n, idx))
            try:
                page = load(ctx)
            except Exception as e:
                print("   載入失敗：%s" % short(e))
                continue
            seen = []

            def on_response(r, seen=seen):
                try:
                    kind = r.request.resource_type
                    if kind in SKIP_KIND or TRACK.search(r.url):
                        return
                    seen.append("%s｜%s｜%s" % (r.status, kind, r.url[:240]))
                except Exception:
                    pass

            page.on("response", on_response)
            try:
                print("   已選產品：%s" % page.evaluate(JS_PICK, idx))
                page.wait_for_timeout(2500)
                print("   --- 選產品後 ---")
                show_section(page, first)
                ch = page.evaluate(JS_CHOICES)
                print("   貨幣選項形式：%s｜%s" % (ch["kind"], "、".join(ch["labels"])))
                for k, i in enumerate(ch["idx"]):
                    got = page.evaluate(JS_APPLY, [ch["kind"], i])
                    page.wait_for_timeout(3000)
                    print("   --- 選貨幣「%s」後（值 %s）---" % (ch["labels"][k], got))
                    show_section(page, first and k == 0)
                    if first and k == 0:
                        show_vue(page, 2500)
                print("   期間的請求 %d 個：" % len(seen))
                for x in seen[:20]:
                    print("      %s" % x)
            except Exception as e:
                print("   操作出錯：%s" % short(e))
            first = False
            page.close()
        browser.close()
    print("\n=== 探查完成 ===")


if __name__ == "__main__":
    main()

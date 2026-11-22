import re
from playwright.sync_api import sync_playwright

URL = "https://www.chubblife.com.hk/zh/customer-service/fulfillment-ratios-of-dividend.html"
KEYS = ["分紅實現率", "總現金價值", "報告年度", "保單年度", "不適用", "N/A",
        "終期紅利", "週年紅利", "周年紅利", "歸原紅利", "%"]
SKIP_KIND = ("image", "font", "stylesheet", "media")
TRACK = re.compile(r"google|doubleclick|facebook|adsrvr|linkedin|yimg|omtrdc|analytics|demdex", re.I)

JS_SELECT = r"""() => {
  const clean = s => (s || '').replace(/\s+/g, ' ').trim();
  const vis = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
  const cut = (e, n) => {
    if (!e) return '';
    const c = e.cloneNode(true);
    c.querySelectorAll('option').forEach((o, i) => { if (i > 1) o.remove(); });
    return c.outerHTML.replace(/\s+/g, ' ').slice(0, n);
  };
  return Array.from(document.querySelectorAll('select')).map(s => {
    let box = s;
    for (let i = 0; i < 3 && box.parentElement; i++) box = box.parentElement;
    return {
      name: s.name || s.id, visible: vis(s),
      attrs: Array.from(s.attributes).map(a => a.name + '=' + a.value.slice(0, 200)),
      options: Array.from(s.options).map(o => ({
        text: clean(o.textContent), value: o.value,
        attrs: Array.from(o.attributes).filter(a => a.name !== 'value')
          .map(a => a.name + '=' + a.value.slice(0, 200))})),
      html: cut(box, 2500),
      next: cut(box.nextElementSibling, 1500)
    };
  });
}"""

JS_SCRIPTS = r"""() => {
  const out = {inline: [], src: []};
  document.querySelectorAll('script').forEach(s => {
    if (s.src) {
      if (!/google|facebook|adsrvr|linkedin|yimg|omtrdc|demdex|doubleclick/i.test(s.src)) out.src.push(s.src);
      return;
    }
    const t = s.textContent || '';
    const re = /fulfil/ig;
    let m, n = 0;
    while ((m = re.exec(t)) && n < 3) {
      out.inline.push(t.slice(Math.max(0, m.index - 200), m.index + 400).replace(/\s+/g, ' '));
      n++;
    }
  });
  out.src = out.src.slice(0, 40);
  out.inline = out.inline.slice(0, 8);
  return out;
}"""

JS_STATE = r"""() => {
  const clean = s => (s || '').replace(/\s+/g, ' ').trim();
  const vis = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
  return {
    text: document.body ? clean(document.body.textContent) : '',
    tables: Array.from(document.querySelectorAll('table')).map(t => ({
      rows: t.rows.length,
      cols: Math.max(0, ...Array.from(t.rows).map(r => r.cells.length)),
      visible: vis(t),
      text: clean(t.textContent).slice(0, 700)})),
    links: Array.from(document.querySelectorAll('a[href]'))
      .map(a => clean(a.textContent).slice(0, 80) + '｜' + a.href)
      .filter(x => /\.(pdf|xlsx?|csv|json)(\?|#|$)/i.test(x)),
    frames: Array.from(document.querySelectorAll('iframe, embed, object'))
      .map(f => f.src || f.data || '')
      .filter(x => x && !/adsrvr|google|facebook|doubleclick/i.test(x)),
    marked: Array.from(document.querySelectorAll('[id*="fulfil" i], [class*="fulfil" i]')).slice(0, 25)
      .map(e => e.tagName + '｜id=' + e.id + '｜class=' + String(e.className).slice(0, 80) +
                '｜可見 ' + vis(e) + '｜' + clean(e.textContent).slice(0, 150))
  };
}"""

JS_PICK = r"""(idx) => {
  const s = document.querySelector('select[name="fulfillment-radios-type"], select#fulfillment-radios-type')
            || document.querySelector('select');
  if (!s || idx >= s.options.length) return null;
  s.selectedIndex = idx;
  s.dispatchEvent(new Event('input', {bubbles: true}));
  s.dispatchEvent(new Event('change', {bubbles: true}));
  return {text: (s.options[idx].textContent || '').replace(/\s+/g, ' ').trim(), value: s.value};
}"""


def short(e):
    return re.sub(r"\s+", " ", str(e))[:300]


def load(ctx):
    page = ctx.new_page()
    page.goto(URL, wait_until="domcontentloaded", timeout=60000)
    try:
        page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:
        pass
    page.wait_for_timeout(3000)
    return page


def state(page):
    try:
        return page.evaluate(JS_STATE)
    except Exception as e:
        print("   讀取頁面狀態失敗：%s" % short(e))
        return None


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False,
                                    args=["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(locale="zh-HK", timezone_id="Asia/Hong_Kong",
                                  viewport={"width": 1366, "height": 900},
                                  accept_downloads=True)
        page = load(ctx)
        selects = page.evaluate(JS_SELECT)
        print("=== 1. 下拉選單（%d 個）===" % len(selects))
        count = 0
        for s in selects:
            print("   名稱 %s｜可見 %s｜選項 %d 個" % (s["name"], s["visible"], len(s["options"])))
            print("   屬性：%s" % "；".join(s["attrs"]))
            for i, o in enumerate(s["options"]):
                print("   [%d] %s｜value=%s｜%s" % (i, o["text"], o["value"][:200], "；".join(o["attrs"])))
            print("   選單外層 HTML：%s" % s["html"])
            print("   選單之後的 HTML：%s" % s["next"])
            count = max(count, len(s["options"]))

        sc = page.evaluate(JS_SCRIPTS)
        print("\n=== 2. 程式碼線索 ===")
        print("   內嵌程式提到 fulfil 的片段 %d 段：" % len(sc["inline"]))
        for x in sc["inline"]:
            print("      %s" % x)
        print("   外部程式檔 %d 個：" % len(sc["src"]))
        for x in sc["src"]:
            print("      %s" % x[:220])

        base = state(page)
        print("\n=== 3. 選產品前 ===")
        if base:
            print("   文字長度 %d｜表格 %d 個｜檔案連結 %d 個" % (
                len(base["text"]), len(base["tables"]), len(base["links"])))
            print("   名稱含 fulfil 的元素 %d 個：" % len(base["marked"]))
            for x in base["marked"]:
                print("      %s" % x)
            for x in base["frames"]:
                print("   內嵌框架：%s" % x[:220])
        page.close()

        picks = []
        for i in (1, 2, count - 1):
            if 0 < i < count and i not in picks:
                picks.append(i)
        for n, idx in enumerate(picks, 1):
            print("\n=== 4.%d 選第 [%d] 項 ===" % (n, idx))
            try:
                page = load(ctx)
            except Exception as e:
                print("   載入失敗：%s" % short(e))
                continue
            seen, popups, downloads = [], [], []

            def on_response(r, seen=seen):
                try:
                    kind = r.request.resource_type
                    if kind in SKIP_KIND or TRACK.search(r.url):
                        return
                    seen.append("%s｜%s｜%s｜%s" % (
                        r.status, kind, r.headers.get("content-type", "")[:40], r.url[:240]))
                except Exception:
                    pass

            page.on("response", on_response)
            page.on("popup", lambda pp, popups=popups: popups.append(pp))
            page.on("download", lambda d, downloads=downloads: downloads.append(
                "%s｜%s" % (d.suggested_filename, d.url[:240])))
            before = state(page)
            del seen[:]
            try:
                got = page.evaluate(JS_PICK, idx)
            except Exception as e:
                print("   選擇時出錯：%s" % short(e))
                got = None
            print("   已選：%s" % got)
            page.wait_for_timeout(6000)
            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass
            print("   目前網址：%s" % page.url)
            print("   彈出視窗 %d 個：" % len(popups))
            for pp in popups:
                try:
                    pp.wait_for_timeout(2000)
                    print("      %s" % pp.url[:240])
                except Exception as e:
                    print("      讀不到網址：%s" % short(e))
            print("   下載 %d 個：" % len(downloads))
            for x in downloads:
                print("      %s" % x)
            print("   選後的請求 %d 個：" % len(seen))
            for x in seen[:40]:
                print("      %s" % x)
            after = state(page)
            if before and after:
                print("   文字長度 %d → %d｜表格 %d → %d｜檔案連結 %d → %d" % (
                    len(before["text"]), len(after["text"]),
                    len(before["tables"]), len(after["tables"]),
                    len(before["links"]), len(after["links"])))
                old = set(t["text"] for t in before["tables"])
                fresh = [t for t in after["tables"] if t["text"] not in old]
                print("   新出現的表格 %d 個：" % len(fresh))
                for t in fresh[:6]:
                    print("      %d 行 × %d 欄｜可見 %s｜%s" % (
                        t["rows"], t["cols"], t["visible"], t["text"]))
                for x in [k for k in after["links"] if k not in before["links"]][:20]:
                    print("   新檔案連結：%s" % x[:300])
                for x in [k for k in after["frames"] if k not in before["frames"]][:10]:
                    print("   新內嵌框架：%s" % x[:240])
                for x in [k for k in after["marked"] if k not in before["marked"]][:15]:
                    print("   有變化的 fulfil 元素：%s" % x)
                had = set(before["text"].split(" "))
                new = " ".join(w for w in after["text"].split(" ") if w not in had)
                print("   新出現的文字（%d 字）：%s" % (len(new), new[:1500]))
                print("   關鍵字次數（選前 → 選後）：%s" % "｜".join(
                    "%s %d→%d" % (k, before["text"].count(k), after["text"].count(k)) for k in KEYS))
            for pp in popups:
                try:
                    pp.close()
                except Exception:
                    pass
            try:
                page.close()
            except Exception:
                pass
        browser.close()
    print("\n=== 探查完成 ===")


if __name__ == "__main__":
    main()

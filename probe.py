import re
from playwright.sync_api import sync_playwright

URLS = [
    "https://www.chubblife.com.hk/zh/customer-service/fulfillment-ratios-of-dividend.html",
    "https://www.chubb.com/hk-zh/customer-service/fulfillment-ratios-of-dividend.html",
]
KEYS = ["分紅實現率", "總現金價值", "報告年度", "保單年度", "不適用", "N/A",
        "終期紅利", "週年紅利", "周年紅利", "美元", "港元"]

JS = r"""() => {
  const clean = s => (s || '').replace(/\s+/g, ' ').trim();
  const vis = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
  const tables = Array.from(document.querySelectorAll('table'));
  return {
    title: document.title,
    text: document.body ? document.body.textContent.replace(/\s+/g, ' ') : '',
    headings: Array.from(document.querySelectorAll('h1,h2,h3,h4,h5'))
      .map(e => e.tagName + ' ' + clean(e.textContent)).filter(s => s.length > 3).slice(0, 80),
    tables: tables.length,
    visible: tables.filter(vis).length,
    sample: tables.slice(0, 8).map(t => ({
      rows: t.rows.length,
      cols: Math.max(0, ...Array.from(t.rows).map(r => r.cells.length)),
      visible: vis(t),
      text: clean(t.textContent).slice(0, 500)})),
    iframes: Array.from(document.querySelectorAll('iframe')).map(f => f.src).slice(0, 10),
    selects: Array.from(document.querySelectorAll('select')).map(s => ({
      name: s.name || s.id,
      options: Array.from(s.options).map(o => clean(o.textContent)).slice(0, 40)})).slice(0, 10),
    links: Array.from(document.querySelectorAll('a[href]'))
      .map(a => ({t: clean(a.textContent).slice(0, 80), h: a.href}))
      .filter(x => /\.(pdf|xlsx?|csv|json)(\?|#|$)/i.test(x.h) ||
                   /實現率|現金價值|紅利|fulfil|dividend/i.test(x.t + ' ' + x.h)).slice(0, 120)
  };
}"""


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False,
                                    args=["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(locale="zh-HK", timezone_id="Asia/Hong_Kong",
                                  viewport={"width": 1366, "height": 900})
        for n, url in enumerate(URLS, 1):
            print("\n########## 頁面 %d：%s" % (n, url))
            page = ctx.new_page()
            seen = []

            def on_response(r, seen=seen):
                try:
                    ct = r.headers.get("content-type", "")
                    kind = r.request.resource_type
                    if kind in ("xhr", "fetch") or "json" in ct or \
                            re.search(r"\.(pdf|json|xlsx?|csv)(\?|$)", r.url, re.I):
                        seen.append("%s｜%s｜%s｜%s" % (r.status, kind, ct[:40], r.url[:220]))
                except Exception:
                    pass

            page.on("response", on_response)
            try:
                resp = page.goto(url, wait_until="domcontentloaded", timeout=60000)
                try:
                    page.wait_for_load_state("networkidle", timeout=20000)
                except Exception:
                    pass
                page.wait_for_timeout(4000)
                d = page.evaluate(JS)
            except Exception as e:
                print("載入失敗：%s" % re.sub(r"\s+", " ", str(e))[:300])
                page.close()
                continue

            print("=== 1. 基本 ===")
            print("   HTTP %s｜最終網址 %s" % (resp.status if resp else "?", page.url))
            print("   標題：%s｜文字長度 %d" % (d["title"], len(d["text"])))
            print("=== 2. 標題層級（%d 個）===" % len(d["headings"]))
            for h in d["headings"]:
                print("   %s" % h[:120])
            print("=== 3. 表格：%d 個，可見 %d 個｜iframe %d 個 ===" % (
                d["tables"], d["visible"], len(d["iframes"])))
            for i, t in enumerate(d["sample"], 1):
                print("   表格 %d｜%d 行 × %d 欄｜可見 %s｜%s" % (
                    i, t["rows"], t["cols"], t["visible"], t["text"]))
            for f in d["iframes"]:
                print("   iframe：%s" % f[:220])
            print("=== 4. 下拉選單（%d 個）===" % len(d["selects"]))
            for s in d["selects"]:
                print("   %s：%s" % (s["name"], "｜".join(s["options"])))
            print("=== 5. 檔案及相關連結（%d 個）===" % len(d["links"]))
            for x in d["links"]:
                print("   %s｜%s" % (x["t"], x["h"][:220]))
            print("=== 6. 背後請求（%d 個）===" % len(seen))
            for s in seen[:80]:
                print("   %s" % s)
            print("=== 7. 關鍵字 ===")
            for k in KEYS:
                hits = [m.start() for m in re.finditer(re.escape(k), d["text"])]
                print("   「%s」%d 次" % (k, len(hits)))
                for pos in hits[:3]:
                    print("      …%s…" % d["text"][max(0, pos - 50):pos + 90])
            page.close()
        browser.close()
    print("\n=== 探查完成 ===")


if __name__ == "__main__":
    main()

import os
import re
import sys
import glob
import requests

API = "https://api.notion.com/v1"
VER = "2025-09-03"
MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def fail(msg):
    print("✗ %s" % msg)
    print("=== 中止：沒有上傳 ===")
    sys.exit(1)


def call(method, path, token, **kw):
    headers = {"Authorization": "Bearer " + token, "Notion-Version": VER}
    r = requests.request(method, API + path, headers=headers, timeout=120, **kw)
    if r.status_code >= 300:
        fail("Notion %s %s 失敗：HTTP %s｜%s" % (method, path, r.status_code, r.text[:400]))
    return r.json()


def main():
    token = os.environ.get("NOTION_TOKEN", "").strip()
    url = os.environ.get("NOTION_DATABASE_URL", "").strip()
    if not token or not url:
        fail("未設定 NOTION_TOKEN 或 NOTION_DATABASE_URL")

    files = glob.glob("chubb_fulfillment_*.xlsx")
    if len(files) != 1:
        fail("找到 %d 個 Excel，應該只有 1 個" % len(files))
    m = re.search(r"chubb_fulfillment_(\d{4})\.xlsx$", files[0])
    if not m:
        fail("檔名讀不到報告年度：%s" % files[0])
    year = m.group(1)
    name = "安達人壽分紅實現率（%s年）自動生成.xlsx" % year

    path = url.split("?")[0].rstrip("/").replace("-", "").lower()
    m = re.search(r"([0-9a-f]{32})$", path)
    if not m:
        fail("NOTION_DATABASE_URL 讀不到資料庫 ID")
    db = call("GET", "/databases/" + m.group(1), token)
    print("=== 1. 資料庫 ===")
    print("   名稱：%s" % "".join(t.get("plain_text", "") for t in db.get("title", [])))
    sources = db.get("data_sources") or []
    if len(sources) != 1:
        fail("資料庫有 %d 個資料來源，應該只有 1 個" % len(sources))
    ds_id = sources[0]["id"]
    ds = call("GET", "/data_sources/" + ds_id, token)
    titles = [k for k, v in (ds.get("properties") or {}).items() if v.get("type") == "title"]
    if len(titles) != 1:
        fail("找不到標題欄位")
    title = titles[0]

    print("=== 2. 檢查是否已有「%s」===" % year)
    q = call("POST", "/data_sources/%s/query" % ds_id, token, json={
        "filter": {"property": title, "title": {"contains": year}}, "page_size": 10})
    if q.get("results"):
        for p in q["results"]:
            print("   已存在：%s" % p.get("url"))
        print("=== 完成：已有 %s 頁面，不重複上傳 ===" % year)
        return

    print("=== 3. 上傳 %s ===" % name)
    up = call("POST", "/file_uploads", token, json={"filename": name, "content_type": MIME})
    with open(files[0], "rb") as fh:
        sent = call("POST", "/file_uploads/%s/send" % up["id"], token,
                    files={"file": ("upload.xlsx", fh, MIME)})
    if sent.get("status") != "uploaded":
        fail("檔案狀態不是 uploaded：%s" % sent.get("status"))

    page = call("POST", "/pages", token, json={
        "parent": {"type": "data_source_id", "data_source_id": ds_id},
        "properties": {title: {"title": [{"text": {"content": year}}]}},
        "children": [{"object": "block", "type": "file",
                      "file": {"type": "file_upload", "file_upload": {"id": up["id"]}}}]})
    print("   已上傳：建立「%s」頁面｜%s" % (year, page.get("url")))
    print("=== 完成 ===")


if __name__ == "__main__":
    main()

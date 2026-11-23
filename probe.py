import io
import re
import urllib.parse

import pdfplumber
import requests

PAGE = "https://www.fubonlife.com.hk/products_philosophy.html"
PDF = "https://www.fubonlife.com.hk/public/pdf/Fubon%20Life%20Fulfilment%20Ratios%20Disclosure_Chi.pdf"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
H = {"User-Agent": UA, "Accept-Language": "zh-HK,zh;q=0.9"}


def clean(s):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", s or "")).strip()


print("=== 1. 監管披露頁 ===")
try:
    r = requests.get(PAGE, headers=H, timeout=60)
    r.encoding = "utf-8"
    print("   HTTP %s｜%d 字元" % (r.status_code, len(r.text)))
    seen = []
    for m in re.finditer(r'<a\b[^>]*href="([^"]+)"[^>]*>(.*?)</a>', r.text, flags=re.S | re.I):
        href, label = m.group(1), clean(m.group(2))
        if ".pdf" in href.lower() or re.search(r"實現率|現金價值|派息率|fulfil", href + label, flags=re.I):
            url = urllib.parse.urljoin(PAGE, href)
            if url not in seen:
                seen.append(url)
                print("   %s｜%s" % (label[:40] or "（無文字）", url))
    print("   相關連結 %d 個" % len(seen))
except Exception as e:
    print("   ✗ 讀取失敗：%s" % e)

print("\n=== 2. 分紅實現率 PDF ===")
r = requests.get(PDF, headers=H, timeout=120)
print("   HTTP %s｜%s｜%d KB｜Last-Modified：%s" % (
    r.status_code, r.headers.get("Content-Type"), len(r.content) // 1024,
    r.headers.get("Last-Modified")))
if r.status_code != 200 or not r.content.startswith(b"%PDF"):
    raise SystemExit("✗ 取不到 PDF")

with pdfplumber.open(io.BytesIO(r.content)) as pdf:
    print("   頁數：%d｜資料：%s" % (len(pdf.pages), pdf.metadata))
    for i, page in enumerate(pdf.pages, 1):
        text = page.extract_text() or ""
        tables = page.extract_tables()
        print("\n--- 第 %d 頁｜文字 %d 字元｜表格 %d 個 ---" % (i, len(text), len(tables)))
        print(text[:1500])
        for k, t in enumerate(tables, 1):
            width = max(len(row) for row in t) if t else 0
            print("   [表格 %d] %d 行 × %d 欄" % (k, len(t), width))
            for row in t[:12]:
                print("      " + "｜".join(clean(str(c)) if c is not None else "∅" for c in row))
            if len(t) > 12:
                print("      …（另有 %d 行）" % (len(t) - 12))
print("\n=== 探查完成：沒有產生 Excel，沒有連接 Notion ===")

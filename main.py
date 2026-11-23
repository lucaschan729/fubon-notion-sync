  import io
  import math
  import os
  import re
  import urllib.parse
  from datetime import datetime, timedelta, timezone

  import pdfplumber
  import requests
  from openpyxl import Workbook
  from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
  from openpyxl.utils import get_column_letter

  PAGE = "https://www.fubonlife.com.hk/products_philosophy.html"
  PDF_FALLBACK = "https://www.fubonlife.com.hk/public/pdf/Fubon%20Life%20Fulfilment%20Ratios%20Disclosure_Chi.pdf"
  H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                     "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
       "Accept-Language": "zh-HK,zh;q=0.9"}
  FOOTER = ("富邦人壽保險（香港）有限公司", "香港太古城", "Fubon Life Insurance", "Suites 701")
  CJK = r"\u3000-\u303f\u4e00-\u9fff\uff00-\uffef"
  KNOWN = ("累積週年紅利及利息", "終期紅利", "特別紅利", "累積利息")
  API = "https://api.notion.com/v1"


  def squash(s):
      return re.sub(r"\s+", "", s or "")


  def name_clean(s):
      s = re.sub(r"\s+", " ", s or "").strip()
      return re.sub(r"(?<=[%s])\s+(?=[%s])" % (CJK, CJK), "", s)


  def body_lines(text):
      out = []
      for line in text.splitlines():
          line = line.strip()
          if line and not line.startswith(FOOTER):
              out.append(line)
      return out


  def find_pdf():
      try:
          r = requests.get(PAGE, headers=H, timeout=60)
          r.encoding = "utf-8"
          print("   HTTP %s｜%s" % (r.status_code, PAGE))
          print("   頁面有否「現金價值」字眼：%s" % ("有" if "現金價值" in r.text else "沒有"))
          m = re.search(r"downloadFile\('([^']*Fulfil[^']*Chi\.pdf)'\)", r.text, flags=re.I)
          if m:
              return urllib.parse.urljoin(PAGE, urllib.parse.quote(m.group(1)))
          print("   ⚠ 監管披露頁找不到分紅實現率 PDF 連結，改用固定網址")
      except Exception as e:
          print("   ⚠ 監管披露頁讀取失敗（%s），改用固定網址" % e)
      return PDF_FALLBACK


  def parse(data):
      rows, years, section, checks, notes = [], None, "", [], []
      with pdfplumber.open(io.BytesIO(data)) as pdf:
          created = re.search(r"D:(\d{4})(\d{2})(\d{2})", (pdf.metadata or {}).get("CreationDate", "") or "")
          texts = [p.extract_text() or "" for p in pdf.pages]
          for i, page in enumerate(pdf.pages):
              text = texts[i]
              if re.search(r"ii\)\s*其他", text):
                  section = "其他"
              elif re.search(r"i\)\s*分紅保單", text):
                  section = "分紅保單"
              n_pct = n_na = n_rows = 0
              for t in page.extract_tables():
                  if not t or max(len(r) for r in t) < 15:
                      continue
                  for raw in t:
                      if raw[0] is None or not squash(raw[0]) or squash(raw[0]).startswith("產品系列"):
                          continue
                      cells = [c for c in raw if c is not None and squash(c)]
                      if len(cells) != 15:
                          raise SystemExit("✗ 第 %d 頁有一行不是 15 格：%s" % (i + 1, cells))
                      vals = []
                      for c in cells[5:]:
                          v = squash(c)
                          if re.fullmatch(r"\d+(\.\d+)?%", v):
                              vals.append(float(v[:-1]))
                              n_pct += 1
                          elif v.startswith("不適用"):
                              vals.append(None)
                              n_na += 1
                          else:
                              raise SystemExit("✗ 第 %d 頁有看不懂的數值：%s" % (i + 1, v))
                      rtype = squash(cells[4])
                      if rtype not in KNOWN:
                          print("   ⚠ 新的紅利類型：%s" % rtype)
                      rows.append({"section": section, "name": name_clean(cells[0]), "ptype": squash(cells[1]),
                                   "launch": squash(cells[2]), "cur": squash(cells[3]), "rtype": rtype, "vals": vals})
                      n_rows += 1
              if n_rows:
                  ys = [int(y) for y in re.findall(r"\((\d{4})\s*年生", text)[:10]]
                  if len(ys) != 10 or (years and ys != years):
                      raise SystemExit("✗ 第 %d 頁的年度表頭不是 10 欄或與前頁不同：%s" % (i + 1, ys))
                  years = ys
                  t_pct = len(re.findall(r"\d+(?:\.\d+)?%", text))
                  t_na = len(re.findall(r"不適用\s*\(\d\)", text))
                  checks.append((i + 1, n_rows, n_pct, t_pct, n_na, t_na))
              elif "註" in text and i > 0:
                  for line in body_lines(text):
                      if re.match(r"^(註|•|\(\d\)|請注意)", line) or not notes:
                          notes.append(line)
                      else:
                          notes[-1] += line
      if not rows:
          raise SystemExit("✗ PDF 內找不到資料行")
      m = re.search(r"(\d{4})\s*報告年度", texts[0])
      if not m:
          raise SystemExit("✗ 找不到報告年度")
      pub = re.search(r"於\s*(\d{4})\s*年\s*(\d{1,2})\s*月公佈", "\n".join(texts))
      intro = "".join(body_lines(texts[0]))
      intro = intro[intro.find("以下"):] if "以下" in intro else intro
      return {"year": int(m.group(1)), "pub": "%s年%s月" % pub.groups() if pub else "未標示",
              "created": "-".join(created.groups()) if created else "未知", "years": years, "rows": rows,
              "checks": checks, "intro": [s + "。" for s in intro.split("。") if s.strip()], "notes": notes}


  def stat(rows):
      vals = [v for r in rows for v in r["vals"] if v is not None]
      return {"products": len({r["name"] for r in rows}), "rows": len(rows), "n": len(vals),
              "avg": sum(vals) / len(vals) if vals else None, "hi": max(vals) if vals else None,
              "lo": min(vals) if vals else None, "low": sum(v < 80 for v in vals),
              "high": sum(v >= 100 for v in vals), "over": sum(v > 100 for v in vals)}


  def pct(v, digits=0):
      return "/" if v is None else ("%.*f%%" % (digits, v))


  def groups(rows):
      keys = []
      for r in rows:
          if (r["section"], r["rtype"]) not in keys:
              keys.append((r["section"], r["rtype"]))
      return [(k, stat([r for r in rows if (r["section"], r["rtype"]) == k])) for k in keys]


  def oddities(d):
      out = []
      for r in d["rows"]:
          m = re.match(r"(\d{4})", r["launch"])
          for k, v in enumerate(r["vals"]):
              if m and v is not None and d["years"][k] < int(m.group(1)):
                  out.append("%s｜%s｜%s推出，但保單年度%d（%d年生效）顯示 %s（照錄官網）"
                             % (r["name"], r["rtype"], r["launch"], k + 1, d["years"][k], pct(v)))
      return out


  def build_excel(d, path, pdf_url, now, odd):
      wb = Workbook()
      ov = wb.active
      ov.title = "概覽"
      ws = wb.create_sheet("%s分紅實現率" % d["year"])
      thin = Side(style="thin", color="BFBFBF")
      box = Border(left=thin, right=thin, top=thin, bottom=thin)
      navy, blue = PatternFill("solid", fgColor="1F4E78"), PatternFill("solid", fgColor="2F75B5")
      grey = PatternFill("solid", fgColor="D9D9D9")
      green, red = PatternFill("solid", fgColor="C6EFCE"), PatternFill("solid", fgColor="FFC7CE")
      white = Font(bold=True, color="FFFFFF")
      center = Alignment(horizontal="center", vertical="center", wrap_text=True)
      left = Alignment(horizontal="left", vertical="center", wrap_text=True)
      title = "富邦人壽（香港）分紅實現率｜%s 報告年度" % d["year"]

      heads = ["产品名称", "产品种类", "推出日期", "货币", "红利类型"] + \
              ["保单年度%d\n(%d年生效)" % (k + 1, y) for k, y in enumerate(d["years"])]
      n = len(heads)
      ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
      c = ws.cell(1, 1, title)
      c.font, c.fill, c.alignment = Font(bold=True, color="FFFFFF", size=14), navy, center
      ws.row_dimensions[1].height = 28
      for j, h in enumerate(heads, 1):
          c = ws.cell(2, j, h)
          c.font, c.fill, c.alignment, c.border = white, navy, center, box
      ws.row_dimensions[2].height = 36
      for j, w in enumerate([34, 16, 13, 8, 22] + [13] * 10, 1):
          ws.column_dimensions[get_column_letter(j)].width = w
      r, sec, prev = 3, None, None
      for row in d["rows"]:
          if row["section"] != sec:
              sec, prev = row["section"], None
              ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n)
              c = ws.cell(r, 1, "【%s】" % sec)
              c.font, c.fill, c.alignment = white, blue, left
              r += 1
          elif row["name"] != prev:
              for j in range(1, n + 1):
                  ws.cell(r, j).fill = grey
              ws.row_dimensions[r].height = 5
              r += 1
          prev = row["name"]
          for j, v in enumerate([row["name"], row["ptype"], row["launch"], row["cur"], row["rtype"]], 1):
              c = ws.cell(r, j, v)
              c.border, c.alignment = box, (left if j == 1 else center)
          ws.cell(r, 1).font = Font(bold=True)
          ws.cell(r, 4).font = Font(bold=True, color="7030A0")
          for k, v in enumerate(row["vals"]):
              c = ws.cell(r, 6 + k)
              c.border, c.alignment = box, center
              if v is None:
                  c.value, c.font = "/", Font(color="808080")
              else:
                  c.value, c.number_format = v / 100, "0%"
                  if v >= 100:
                      c.fill = green
                  elif v < 80:
                      c.fill = red
          ws.row_dimensions[r].height = 32 if len(row["name"]) > 16 else 20
          r += 1
      ws.freeze_panes = "B3"

      ov.column_dimensions["A"].width = 24
      for j in range(2, 11):
          ov.column_dimensions[get_column_letter(j)].width = 13
      ov.merge_cells("A1:J1")
      c = ov.cell(1, 1, title)
      c.font, c.fill, c.alignment = Font(bold=True, color="FFFFFF", size=14), navy, center
      ov.row_dimensions[1].height = 28
      r = 2

      def kv(k, v):
          nonlocal r
          ov.cell(r, 1, k).font = Font(bold=True)
          ov.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
          ov.cell(r, 2, v).alignment = left
          ov.row_dimensions[r].height = 18 * max(1, math.ceil(len(str(v)) / 55))
          r += 1

      def long(text, bold=False):
          nonlocal r
          ov.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
          c = ov.cell(r, 1, text)
          c.alignment, c.font = left, Font(bold=bold)
          ov.row_dimensions[r].height = 18 * max(1, math.ceil(len(text) / 62)) + 2
          r += 1

      total = stat(d["rows"])
      kv("官網公佈日期", d["pub"])
      kv("PDF 建立日期", d["created"])
      kv("抓取時間（香港）", now)
      kv("來源頁", PAGE)
      kv("資料檔", pdf_url)
      kv("產品數目", "%d 個" % total["products"])
      kv("核對", "%d 頁表格的數值個數與同頁文字一致" % len(d["checks"]))
      kv("圖例", "綠底：100% 或以上｜紅底：低於 80%｜/：官網顯示「不適用」")
      if odd:
          kv("留意", "；".join(odd))
      r += 1
      for j, h in enumerate(["類別", "紅利類型", "產品數", "行數", "數值個數", "平均", "最高", "最低",
                             "低於 80%", "100% 或以上"], 1):
          c = ov.cell(r, j, h)
          c.font, c.fill, c.alignment, c.border = white, navy, center, box
      r += 1
      for (s, t), st in groups(d["rows"]) + [(("合計", ""), total)]:
          line = [s, t, st["products"], st["rows"], st["n"], pct(st["avg"], 2), pct(st["hi"]), pct(st["lo"]),
                  st["low"], st["high"]]
          for j, v in enumerate(line, 1):
              c = ov.cell(r, j, v)
              c.border, c.alignment = box, center
              if s == "合計":
                  c.font = Font(bold=True)
          ov.row_dimensions[r].height = 32
          r += 1
      r += 1
      long("官網說明", True)
      for s in d["intro"]:
          long(s)
      long("官網備註", True)
      for s in d["notes"]:
          long(s)
      wb.save(path)
      return [ov.title, ws.title]


  def notion(method, path, token, **kw):
      h = {"Authorization": "Bearer " + token, "Notion-Version": "2025-09-03"}
      r = requests.request(method, API + path, headers=h, timeout=120, **kw)
      if r.status_code >= 300:
          raise SystemExit("   ✗ 失敗：Notion %s %s｜HTTP %s｜%s" % (method, path.split("?")[0], r.status_code, r.text[:300]))
      return r.json()


  def database_id():
      raw = os.environ.get("NOTION_DATABASE_ID") or os.environ.get("NOTION_DATABASE_URL") or ""
      raw = raw.strip().split("?")[0].rstrip("/").replace("-", "")[-32:]
      return raw if re.fullmatch(r"[0-9a-fA-F]{32}", raw) else ""


  def upload(path, year):
      token, did = os.environ.get("NOTION_TOKEN", ""), database_id()
      if not token or not did:
          raise SystemExit("   ✗ 未設定 NOTION_TOKEN，或 NOTION_DATABASE_ID／NOTION_DATABASE_URL 格式不對")
      db = notion("GET", "/databases/" + did, token)
      print("   資料庫：%s" % "".join(t.get("plain_text", "") for t in db.get("title", [])))
      sources = db.get("data_sources") or []
      if not sources:
          raise SystemExit("   ✗ 資料庫沒有 data source")
      ds = sources[0]["id"]
      props = notion("GET", "/data_sources/" + ds, token).get("properties", {})
      tname = next((k for k, v in props.items() if v.get("type") == "title"), None)
      if not tname:
          raise SystemExit("   ✗ 找不到標題欄")
      title = str(year)
      res = notion("POST", "/data_sources/%s/query" % ds, token,
                   json={"filter": {"property": tname, "title": {"equals": title}}, "page_size": 5})
      if res.get("results"):
          pid = res["results"][0]["id"]
          print("   使用現有頁面： %s" % title)
      else:
          pid = notion("POST", "/pages", token, json={
              "parent": {"type": "data_source_id", "data_source_id": ds},
              "properties": {tname: {"title": [{"text": {"content": title}}]}}})["id"]
          print("   已建立頁面：%s" % title)
      old, cursor = [], None
      while True:
          q = "/blocks/%s/children?page_size=100" % pid + ("&start_cursor=" + cursor if cursor else "")
          res = notion("GET", q, token)
          for b in res.get("results", []):
              if b.get("type") == "file" and "自動生成" in (b["file"].get("name") or ""):
                  old.append(b["id"])
          if not res.get("has_more"):
              break
          cursor = res.get("next_cursor")
      fname = "富邦人壽分紅實現率（%s年）自動生成.xlsx" % year
      ct = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      fu = notion("POST", "/file_uploads", token, json={"mode": "single_part", "filename": fname, "content_type": ct})
      with open(path, "rb") as f:
          sent = notion("POST", "/file_uploads/%s/send" % fu["id"], token, files={"file": ("fubon.xlsx", f, ct)})
      if sent.get("status") != "uploaded":
          raise SystemExit("   ✗ 失敗：檔案狀態為 %s" % sent.get("status"))
      notion("PATCH", "/blocks/%s/children" % pid, token, json={"children": [{
          "object": "block", "type": "file",
          "file": {"type": "file_upload", "file_upload": {"id": fu["id"]}}}]})
      for b in old:
          notion("DELETE", "/blocks/" + b, token)
      print("   已上傳：%s（刪除舊的自動生成檔案 %d 個）" % (fname, len(old)))


  def main():
      print("=== 1. 讀取官方頁 ===")
      pdf_url = find_pdf()
      r = requests.get(pdf_url, headers=H, timeout=120)
      print("   資料檔：HTTP %s｜%d KB｜%s" % (r.status_code, len(r.content) // 1024, pdf_url))
      if r.status_code != 200 or not r.content.startswith(b"%PDF"):
          raise SystemExit("✗ 取不到 PDF")
      d = parse(r.content)

      print("\n=== 2. 表格與文字互相核對 ===")
      bad = 0
      for p, n_rows, n_pct, t_pct, n_na, t_na in d["checks"]:
          ok = n_pct == t_pct and n_na == t_na
          bad += not ok
          print("   第 %d 頁｜%d 行｜數值 %d／文字 %d｜不適用 %d／文字 %d｜%s"
                % (p, n_rows, n_pct, t_pct, n_na, t_na, "一致" if ok else "✗ 不一致"))
      if bad:
          raise SystemExit("✗ 有 %d 頁表格與文字不一致，停止" % bad)

      print("\n=== 3. 讀取結果 ===")
      total = stat(d["rows"])
      print("   報告年度 %s｜官網公佈 %s｜PDF 建立 %s｜產品 %d 個" % (d["year"], d["pub"], d["created"], total["products"]))
      print("   官網年度欄：" + "｜".join("保單年度%d(%d)" % (k + 1, y) for k, y in enumerate(d["years"])))
      for (s, t), st in groups(d["rows"]):
          print("   %s｜%s：產品 %d 個｜%d 行｜數值 %d 個｜平均 %s｜最高 %s｜最低 %s｜低於 80%% 的 %d 個"
                % (s, t, st["products"], st["rows"], st["n"], pct(st["avg"], 2), pct(st["hi"]), pct(st["lo"]), st["low"]))
      print("   合計：%d 行｜數值 %d 個｜平均 %s｜最高 %s｜最低 %s｜低於 80%% 的 %d 個｜100%% 或以上 %d 個｜高於 100%% 的 %d 個"
            % (total["rows"], total["n"], pct(total["avg"], 2), pct(total["hi"]), pct(total["lo"]),
               total["low"], total["high"], total["over"]))
      print("   有數值的行：")
      for row in d["rows"]:
          if any(v is not None for v in row["vals"]):
              print("      %s｜%s｜%s｜%s" % (row["name"], row["cur"], row["rtype"], " ".join(pct(v) for v in row["vals"])))
      odd = oddities(d)
      for o in odd:
          print("   ⚠ 留意：" + o)

      print("\n=== 4. 官網說明及備註 ===")
      for s in d["intro"] + d["notes"]:
          print("   " + s)

      print("\n=== 5. Excel ===")
      now = datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d %H:%M")
      path = "fubon_fulfillment_%s.xlsx" % d["year"]
      sheets = build_excel(d, path, pdf_url, now, odd)
      print("   %s｜%d KB｜工作表：%s" % (path, os.path.getsize(path) // 1024, "、".join(sheets)))

      print("\n=== 6. 上傳到 Notion ===")
      if os.environ.get("UPLOAD") == "1":
          upload(path, d["year"])
      else:
          print("   試跑：沒有上傳")
      print("完成")


  if __name__ == "__main__":
      main()

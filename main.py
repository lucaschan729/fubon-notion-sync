import collections
import datetime
import io
import math
import os
import re
import urllib.parse

import pdfplumber
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties

TOKEN = os.environ.get("NOTION_TOKEN", "").strip()
API = "https://api.notion.com/v1"
VERSION = "2025-09-03"
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Notion-Version": VERSION,
    "Content-Type": "application/json",
}
PAGE = "https://www.fubonlife.com.hk/products_philosophy.html"
PDF_FALLBACK = "https://www.fubonlife.com.hk/public/pdf/Fubon%20Life%20Fulfilment%20Ratios%20Disclosure_Chi.pdf"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
H = {"User-Agent": UA, "Accept-Language": "zh-HK,zh;q=0.9"}
FOOTER = ("富邦人壽保險（香港）有限公司", "香港太古城", "Fubon Life Insurance", "Suites 701")
CJK = r"\u3000-\u303f\u4e00-\u9fff\uff00-\uffef"
KNOWN = ("累積週年紅利及利息", "終期紅利", "特別紅利", "累積利息")
PREFERRED = ["累積週年紅利及利息", "累積利息", "終期紅利", "特別紅利"]
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
AUTO = "自動生成"
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
SECTION_FONT = Font(bold=True, size=12, color="FF" + NAVY_HEX)
LINE = Side(style="thin", color="FFD0D7E2")
GROUP = Side(style="thin", color="FF7F9BBF")
BORDER = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
CENTER_FIT = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
LEFT_FIT = Alignment(horizontal="left", vertical="center", shrink_to_fit=True, indent=1)


# ---------- 讀取富邦 PDF ----------
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


def oddities(d):
    out = []
    for r in d["rows"]:
        m = re.match(r"(\d{4})", r["launch"])
        for k, v in enumerate(r["vals"]):
            if m and v is not None and d["years"][k] < int(m.group(1)):
                out.append("%s｜%s｜%s推出，但保單年度%d（%d年生效）顯示 %.0f%%（照錄官網）"
                           % (r["name"], r["rtype"], r["launch"], k + 1, d["years"][k], v))
    return out


# ---------- 平均值 ----------
def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def mean_or_slash(nums):
    return sum(nums) / len(nums) if nums else "/"


def numbers(rows, n_fixed):
    return [v for r in rows for v in r[n_fixed:] if is_num(v)]


def make_header(cols, fixed):
    return list(fixed) + [f"保单年度{i + 1}\n{c}" for i, c in enumerate(cols)]


def build_summary(n_fixed, cols, groups):
    n = len(cols)
    width = n_fixed + n
    head = ["所有年份平均值"] + [f"持滿{i + 1}年平均值" for i in range(1, n)]
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
    return {"head": head, "cum": cum, "year_head": list(cols), "year": [year_line]}


def fmt_pct(v):
    return f"{v * 100:.2f}" if is_num(v) else "/"


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
    """100% 或以上綠底；低於 80% 紅底；沒有數字用淺灰字。"""
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
    if "年金" in str(cat):
        return Font(color="FFC55A11")
    return Font(color="FF2F5597")


def units(s):
    return sum(2 if ord(ch) > 255 else 1 for ch in str(s))


def fit_width(values, low, cap):
    longest = max([units(v) for v in values] or [0])
    return min(max(math.ceil(longest * 1.1) + 4, low), cap)


def add_sheet(wb, title, heading, notes, header, rows, summary, head_label, n_fixed, tab):
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
        for c, v in enumerate(row, 1):
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
                    cell.number_format = "0%"
                heat(cell, v, WHITE)
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
                  "「當年平均值」＝該年度一欄所有數字的平均。平均值為固定數字，不會隨篩選變動。")
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


def add_overview(wb, year, pub, today, items, odd):
    """items：[(簡稱, 工作表名稱, 資料行, 固定欄數)]"""
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

    def bullets(r, lines):
        for t in lines:
            r += 1
            cell = ws.cell(row=r, column=1, value="・" + t)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
            cell.font = NOTE_FONT
            cell.alignment = LEFT_FIT
            ws.row_dimensions[r].height = ROW_H
        return r

    banner(1, f"富邦人壽分紅實現率概覽（報告年度 {year}）", TITLE_FONT, DARK, 40)
    banner(2, f"資料來源：富邦人壽（香港）網站｜官網公佈：{pub}｜更新日期：{today}", NOTE_FONT, PALE, ROW_H)

    r = 4
    section(r, "資料概況")
    r += 1
    table_head(r, ["工作表", "產品／版本數", "行數", "有數字的格數", "整體平均值", "低於 80% 的比例"])
    for _, title, rows, n_fixed in items:
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

    r += 2
    section(r, "按產品種類的整體平均值")
    r += 1
    table_head(r, ["產品種類"] + [short for short, _, _, _ in items])
    cats = []
    for _, _, rows, _ in items:
        for row in rows:
            if row[1] not in cats:
                cats.append(row[1])
    for cat in cats:
        r += 1
        cell = ws.cell(row=r, column=1, value=cat or "（未分類）")
        cell.border = BORDER
        cell.alignment = LEFT_FIT
        cell.font = Font(bold=True, color=cat_font(cat).color)
        for k, (_, _, rows, n_fixed) in enumerate(items):
            v = mean_or_slash(numbers([row for row in rows if row[1] == cat], n_fixed))
            c2 = ws.cell(row=r, column=k + 2, value=v)
            c2.border = BORDER
            c2.alignment = CENTER_FIT
            c2.number_format = "0.00%"
            heat(c2, v, WHITE)
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "顏色說明")
    samples = [(1.0, "100% 或以上：與銷售時的說明一致或更好"),
               (0.79, "低於 80%：明顯低於銷售時的說明"),
               ("/", "沒有數字：官網顯示「不適用」")]
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
    r = bullets(r, ["點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
                    "工作表第 3 行有篩選箭嘴，可按產品種類、貨幣或紅利類型篩選。",
                    "向右或向下捲動時，標題列和產品名稱會固定不動。",
                    "平均值摘要在工作表最底部。",
                    "手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。"])
    if odd:
        r += 2
        section(r, "官網資料留意")
        bullets(r, odd)


# ---------- Notion（沿用保誠） ----------
def check(r, what):
    if r.status_code >= 300:
        print(f"{what} 失敗：{r.status_code} {r.text[:400]}")
        raise SystemExit(1)
    return r.json()


def target_id():
    raw = (os.environ.get("NOTION_DATABASE_ID") or os.environ.get("NOTION_DATABASE_URL") or "").strip()
    return raw.split("?")[0].rstrip("/").replace("-", "")[-32:]


def resolve_source(raw):
    m = re.search(r"[0-9a-fA-F]{32}", raw.replace("-", ""))
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
    print(f"找不到富邦的年份資料庫（狀態 {status}）。401＝金鑰不正確；"
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
    print("=== 1. 讀取官方頁 ===")
    pdf_url = find_pdf()
    r = requests.get(pdf_url, headers=H, timeout=120)
    print("   資料檔：HTTP %s｜%d KB｜%s" % (r.status_code, len(r.content) // 1024, pdf_url))
    if r.status_code != 200 or not r.content.startswith(b"%PDF"):
        raise SystemExit("✗ 取不到 PDF")
    d = parse(r.content)

    print()
    print("=== 2. 表格與文字互相核對 ===")
    bad = 0
    for p, n_rows, n_pct, t_pct, n_na, t_na in d["checks"]:
        ok = n_pct == t_pct and n_na == t_na
        bad += not ok
        print("   第 %d 頁｜%d 行｜數值 %d／文字 %d｜不適用 %d／文字 %d｜%s"
              % (p, n_rows, n_pct, t_pct, n_na, t_na, "一致" if ok else "✗ 不一致"))
    if bad:
        raise SystemExit("✗ 有 %d 頁表格與文字不一致，停止" % bad)

    year = str(d["year"])
    cols = [str(y) for y in d["years"]]
    rows = [[x["name"], x["ptype"], x["cur"], x["rtype"]]
            + [round(v / 100, 4) if v is not None else "/" for v in x["vals"]] for x in d["rows"]]
    nums = numbers(rows, 4)
    odd = oddities(d)

    print()
    print("=== 3. 讀取結果 ===")
    print(f"   報告年度 {year}｜官網公佈 {d['pub']}｜PDF 建立 {d['created']}")
    print(f"   產品 {len({r0[0] for r0 in rows})} 個｜{len(rows)} 行｜數值 {len(nums)} 個"
          f"｜平均 {fmt_pct(mean_or_slash(nums))}%｜最高 {max(nums) * 100:.0f}%｜最低 {min(nums) * 100:.0f}%"
          f"｜低於 80% 的 {sum(v < LOW_LIMIT for v in nums)} 個｜100% 或以上 {sum(v >= HIGH_LIMIT for v in nums)} 個")
    print("   年度欄：", " ".join(cols))
    types = collections.Counter(r0[3] for r0 in rows)
    print("   紅利類型：", "｜".join(f"{k} {v}" for k, v in types.items()))
    for cat in dict.fromkeys(r0[1] for r0 in rows):
        sub = [r0 for r0 in rows if r0[1] == cat]
        print(f"   產品種類 {cat}：{len(sub)} 行｜數值 {len(numbers(sub, 4))} 個｜平均 {fmt_pct(mean_or_slash(numbers(sub, 4)))}")
    for o in odd:
        print("   ⚠ 留意：" + o)
    if len(cols) < 5 or len(rows) < 20 or not nums:
        raise SystemExit("讀到的資料太少，富邦的 PDF 可能改了版面，暫不更新")

    order = [t for t in PREFERRED if t in types] + [t for t in types if t not in PREFERRED]
    groups = [("所有分紅", rows)]
    for t in order:
        sub = [r0 for r0 in rows if r0[3] == t]
        if numbers(sub, 4):
            groups.append((t, sub))
    summary = build_summary(4, cols, groups)
    print("   平均值摘要（持滿 1 至 10 年）：")
    for line in summary["cum"]:
        print("     ", line[0], " ".join(fmt_pct(v) for v in line[1:]))
    print("   當年平均值：", " ".join(fmt_pct(v) for v in summary["year"][0][1:]))

    print()
    print("=== 4. 官網說明及備註（只在日誌顯示） ===")
    for s in d["intro"] + d["notes"]:
        print("   " + s)

    print()
    print("=== 5. Excel ===")
    today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d")
    notes = ("分紅實現率＝實際累積的非保證紅利或利息 ÷ 銷售時所示的相應累積數值。"
             "100% 即與銷售時一致，高於 100% 即比銷售時多。\n"
             "綠色＝100% 或以上；紅色＝低於 80%；「/」＝官網顯示「不適用」"
             "（沒有仍然生效的保單、銷售時所示金額為零，或年內沒有保單終止）。\n"
             f"資料來源：富邦人壽（香港）網站（報告年度 {year}，{d['pub']}公佈）｜更新日期：{today}")
    title = f"{year}分紅實現率"
    wb = Workbook()
    wb.remove(wb.active)
    add_sheet(wb, title, f"富邦人壽分紅實現率（報告年度 {year}）", notes,
              make_header(cols, ["产品名称", "产品种类", "货币", "红利类型"]), rows, summary,
              "分紅類型", 4, "2E75B6")
    add_overview(wb, year, d["pub"], today, [("分紅實現率", title, rows, 4)], odd)
    buf = io.BytesIO()
    wb.save(buf)
    content = buf.getvalue()
    path = f"fubon_fulfillment_{year}.xlsx"
    with open(path, "wb") as f:
        f.write(content)
    print(f"   {path}｜{len(content) // 1024} KB｜工作表：", "、".join(wb.sheetnames))

    print()
    print("=== 6. 上傳到 Notion ===")
    if os.environ.get("UPLOAD") != "1":
        print("   試跑：沒有上傳")
        print("完成")
        return
    target = target_id()
    if not TOKEN or not target:
        raise SystemExit("未設定 NOTION_TOKEN 或 NOTION_DATABASE_ID")
    source = resolve_source(target)
    page = next((p for p in query_pages(source) if year in page_title(p)), None)
    if page is None:
        body = {"parent": {"type": "data_source_id", "data_source_id": source},
                "properties": {"title": {"title": [{"text": {"content": year}}]}}}
        page = check(requests.post(f"{API}/pages", headers=HEADERS, json=body, timeout=30), "建立年度頁面")
        print(f"已建立頁面：{year}")
    else:
        print("使用現有頁面：", page_title(page))
    upload_file(page["id"], f"富邦人壽分紅實現率（{year}年）{AUTO}.xlsx", content)
    print("完成")


if __name__ == "__main__":
    main()

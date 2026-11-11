# -*- coding: utf-8 -*-
"""
在 GVOexcel 原檔裡新增（或重建）「All」工作表

- 把「史跡」到「天文」各分頁的資料接成一張表，表格樣式和「史跡」相同
- þ / 2 / 3 / 4 / 5 五個角色欄是公式，自動帶出各分類頁的打勾
  → 打勾照舊在各分類頁做，All 頁只負責顯示，不用管理兩處
- 儲存格格式、任務欄的超連結、註解、接任地點的顏色都一併帶過來
- 直接改 xlsx 內部 XML，其他分頁（圖片、註解、超連結、格式化條件）原封不動
- 前置發現自動判斷：C1 選「查看角色」(1～5)，前置發現依該角色狀態上色
  （綠＝已完成、紅＝未完成），「可接」欄：✔＝前置都完成、?＝有前置無法判斷
- 開啟工作表保護：表頭列與 þ～5 欄鎖定，其他欄可以修改；篩選、調整格式照常可用
  （不設密碼，要排序時「校閱 → 取消保護工作表」即可）
- 原檔先備份成「檔名.bak.xlsx」

用法：
  把 GVOexcel_xxx.xlsx 拖到「加入All頁.bat」上
  或 python gvo_add_all_sheet.py "GVOexcel_xxx.xlsx"
不需要安裝任何套件（只用 Python 內建模組）
"""
import difflib
import html
import os
import re
import shutil
import sys
import zipfile
from xml.sax.saxutils import escape

FIRST_SHEET = "史跡"
LAST_SHEET = "天文"
ALL_SHEET = "All"
ALL_TABLE = "All類"
STYLE_SHEET = "史跡"          # 表格樣式、欄寬、表頭格式參考這一頁

CHECK_COLS = ["þ", "2", "3", "4", "5"]
PRE_COLS = ["前置發現", "前置發現2", "前置發現3"]
HELPER_COLS = ["前1", "前2", "前3"]       # 隱藏欄：前置判斷結果 1=完成 0=未完成 ?=無法判斷
OUT_HEADERS = [
    "★", "發現物", "視", "考", "宗", "任務／地圖",
    "þ", "2", "3", "4", "5", "可接",
    "接任地點", "發現地點", *PRE_COLS,
    "備註", "ID", "分類", "版本", *HELPER_COLS,
]
# 前置發現欄的已知筆誤（括號前的名稱 → 表裡實際的名稱）。自動近似比對抓不到的寫這裡
ALIAS_MAP = {
    "玻璃首飾": "玻璃飾品",
    "塔德拉爾特阿卡庫斯岩畫": "塔德拉爾特阿卡庫斯石洞壁畫",
    "兩張臉": "二張臉",
    "雲臺山": "雲台山",
    "群雄割據的故事": "群雄割據的物語",
    "公會最初的工作": "公會最初的工作/聖母的彩色玻璃",
    "鯨魚的生態調查": "鯨魚的生態調查(生物1)",
    "變色寶石的礦脈的地圖": "紫翠玉貓眼石",
    "變化的聖人俄羅斯皇帝徹底調查": "俄羅斯皇帝徹底調查",
}
# 近似比對如果配錯了，把前置欄的原文字（括號前的部分）加到這裡，就會維持 ?
NO_FUZZY = set()
# 含這些字的前置欄是技能／取得方式的說明，不是前置任務，不列入判斷（同「索菲亞」）
SKIP_WORDS = ["索菲亞", "學術技能", "學術協會技能", "駱駝隊商"]
HDR = 2                 # 表頭在第 2 列；第 1 列放「查看角色」
SELECT_CELL = "C1"
SELECT_CELL_ABS = "$C$1"
# 前置發現也可能是這些分頁裡的東西（分頁, [用來比對的欄]）
EXTRA_SHEETS = [("港口", ["港口˙村落"]), ("氣象現象", ["氣象現象"]),
                ("傳說", ["傳說", "任務名稱"])]
# All 欄 → 「史跡」頁對應欄（用來抓表頭格式、欄寬、預設格式）
STYLE_MAP = {
    "★": "★", "發現物": "史跡", "視": "視", "考": "考", "宗": "宗",
    "任務／地圖": "任務／地圖", "þ": "þ", "2": "2", "3": "3", "4": "4", "5": "5",
    "接任地點": "接任地點", "發現地點": "發現地點", "前置發現": "前置發現",
    "前置發現2": "前置發現2", "前置發現3": "前置發現3", "備註": "備註",
    "ID": "ID", "分類": "版本", "版本": "版本", "可接": "★",
    "前1": "ID", "前2": "ID", "前3": "ID",
}
EXTRA_WIDTH = {"分類": 10.0, "版本": 8.5, "可接": 5.5, "前1": 4, "前2": 4, "前3": 4}

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT_SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"


# ───────────────────────── 小工具 ─────────────────────────

def attrs(tag_text):
    """把 <tag a="1" b="2"> 的屬性抓成 dict"""
    return dict(re.findall(r'([\w:]+)="([^"]*)"', tag_text))


def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def num_to_col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


def expand_range(ref):
    if ":" not in ref:
        return [ref]
    a, b = ref.split(":")
    c1, r1 = split_ref(a)
    c2, r2 = split_ref(b)
    return [f"{num_to_col(c)}{r}"
            for r in range(r1, r2 + 1)
            for c in range(col_to_num(c1), col_to_num(c2) + 1)]


def resolve(base_dir, target):
    """rels 裡的 Target 轉成 zip 內路徑"""
    if target.startswith("/"):
        return target[1:]
    parts = (base_dir + "/" + target).split("/")
    out = []
    for p in parts:
        if p == "..":
            out.pop()
        elif p and p != ".":
            out.append(p)
    return "/".join(out)


def rels_path(part):
    d, f = part.rsplit("/", 1)
    return f"{d}/_rels/{f}.rels"


def parse_rels(text):
    rels = {}
    for tag in re.findall(r"<Relationship [^>]*?/>", text):
        a = attrs(tag)
        rels[a["Id"]] = a
    return rels


def sref_name(name):
    """結構化參照裡的欄名，特殊字元要加 ' 跳脫"""
    return re.sub(r"(['#\[\]])", r"'\1", name)


# ───────────────────────── 讀檔 ─────────────────────────

class Book:
    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.files = {n: z.read(n) for n in z.namelist()}
            self.order = z.namelist()
        self.wb = self.text("xl/workbook.xml")
        self.wb_rels = parse_rels(self.text("xl/_rels/workbook.xml.rels"))
        self.sheets = []   # [(name, sheetId, rId, tag)]
        # <sheet .../> 或 <sheet ...></sheet>，也接受 x:sheet 這種加前綴的寫法
        for tag in re.findall(r"<(?:\w+:)?sheet\b[^>]*?(?:/>|>\s*</(?:\w+:)?sheet>)", self.wb):
            a = attrs(tag)
            rid = next((v for k, v in a.items() if k.endswith(":id")), None)
            if "name" in a and rid:
                self.sheets.append((unescape_attr(a["name"]).strip(), a.get("sheetId", "0"),
                                    rid, tag))
        self.shared = self._load_shared()

    def text(self, name):
        return normalize_xml(self.files[name].decode("utf-8-sig"))

    def put(self, name, text):
        if name not in self.files:
            self.order.append(name)
        self.files[name] = text.encode("utf-8")

    def remove(self, name):
        self.files.pop(name, None)
        if name in self.order:
            self.order.remove(name)

    def sheet_part(self, rid):
        return resolve("xl", self.wb_rels[rid]["Target"])

    def _load_shared(self):
        if "xl/sharedStrings.xml" not in self.files:
            return []
        out = []
        for si in re.findall(r"<si>(.*?)</si>", self.text("xl/sharedStrings.xml"), re.S):
            # 只取正文 <t>，忽略注音 <rPh>
            si = re.sub(r"<rPh.*?</rPh>", "", si, flags=re.S)
            out.append(unescape_text("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))))
        return out

    def save(self, path):
        tmp = path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for n in self.order:
                z.writestr(n, self.files[n])
        os.replace(tmp, path)


def normalize_xml(text):
    """
    有些程式（例如用 Open XML SDK 存檔）會寫成 <x:sheet>、<x:c> 這種加前綴的標籤，
    意思和 <sheet>、<c> 完全一樣。這裡統一拿掉主命名空間的前綴，後面的處理才認得。
    """
    m = re.search(r'xmlns:(\w+)="' + re.escape(NS_MAIN) + '"', text[:5000])
    if not m:
        return text
    p = m.group(1)
    root = re.search(r"<(?!\?|!)([\w:]+)", text)
    if not root or not root.group(1).startswith(p + ":"):
        return text
    if re.search(r'\sxmlns="', text):      # 文件內另有預設命名空間時不動，避免改錯
        return text
    text = re.sub(rf"<(/?){p}:", r"<\1", text)
    return text.replace(f'xmlns:{p}="{NS_MAIN}"', f'xmlns="{NS_MAIN}"', 1)


def unescape_text(s):
    return html.unescape(s)      # 也處理 &#21490; 這類數字編碼


unescape_attr = unescape_text


class Cell:
    __slots__ = ("attrs", "inner")

    def __init__(self, attr_text, inner):
        self.attrs = attr_text      # 例如 ' s="19" t="s"'（不含 r）
        self.inner = inner or ""    # <v>..</v> / <is>..</is> / <f>..</f><v>..</v>

    @property
    def style(self):
        m = re.search(r'\bs="(\d+)"', self.attrs)
        return m.group(1) if m else None

    def value(self, shared):
        t = re.search(r'\bt="(\w+)"', self.attrs)
        t = t.group(1) if t else "n"
        if t == "inlineStr":
            return unescape_text("".join(re.findall(r"<t[^>]*>(.*?)</t>", self.inner, re.S)))
        v = re.search(r"<v>(.*?)</v>", self.inner, re.S)
        if not v:
            return None
        v = unescape_text(v.group(1))
        if t == "s":
            return shared[int(v)]
        return v

    def has_formula(self):
        return "<f" in self.inner


CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)


def parse_cells(sheet_xml):
    data = re.search(r"<sheetData>(.*)</sheetData>", sheet_xml, re.S)
    cells = {}
    if data:
        for col, row, at, inner in CELL_RE.findall(data.group(1)):
            at = re.sub(r'\s*\br="[^"]*"', "", at)
            cells[(int(row), col)] = Cell(at, inner)
    return cells


class Category:
    """一個分類分頁（史跡、宗教建築…）"""

    def __init__(self, book, name, rid):
        self.name = name
        self.part = book.sheet_part(rid)
        self.xml = book.text(self.part)
        rp = rels_path(self.part)
        self.rels = parse_rels(book.text(rp)) if rp in book.files else {}
        base = self.part.rsplit("/", 1)[0]

        # 表格
        tp = re.search(r'<tablePart r:id="(\w+)"', self.xml)
        if not tp:
            raise ValueError(f"「{name}」沒有 Excel 表格")
        self.table_part = resolve(base, self.rels[tp.group(1)]["Target"])
        txml = book.text(self.table_part)
        self.table_xml = txml
        ta = attrs(re.search(r"<table [^>]*>", txml).group(0))
        # 結構化參照用的是 displayName；有些程式（OnlyOffice）存檔時不寫 name
        self.table = unescape_attr(ta.get("displayName") or ta.get("name") or "")
        if not self.table:
            raise ValueError(f"「{name}」的表格沒有名稱")
        c1, r1 = split_ref(ta["ref"].split(":")[0])
        c2, r2 = split_ref(ta["ref"].split(":")[1])
        totals = int(ta.get("totalsRowCount", "0"))
        self.first_row, self.last_row = r1 + 1, r2 - totals
        self.header_row = r1
        names = [unescape_attr(attrs(t)["name"])
                 for t in re.findall(r"<tableColumn [^>]*?>", txml)]
        self.col_of = {}            # 欄名 → 欄字母
        for i, n in enumerate(names):
            self.col_of.setdefault(n.strip(), num_to_col(col_to_num(c1) + i))
        self.names = [n.strip() for n in names]

        self.cells = parse_cells(self.xml)

        # 超連結
        self.links = {}
        for tag in re.findall(r"<hyperlink [^>]*?/>", self.xml):
            a = attrs(tag)
            for ref in expand_range(a["ref"]):
                self.links[ref] = a

        # 註解
        self.comments = {}
        self.vml_shapes = {}
        for rel in self.rels.values():
            typ = rel["Type"].rsplit("/", 1)[-1]
            if typ == "comments":
                cx = book.text(resolve(base, rel["Target"]))
                authors = [unescape_text(a) for a in
                           re.findall(r"<author>(.*?)</author>", cx, re.S)]
                for ca, body in re.findall(r"<comment ([^>]*)>(.*?)</comment>", cx, re.S):
                    a = attrs(ca)
                    au = authors[int(a.get("authorId", "0"))] if authors else ""
                    self.comments[a["ref"]] = (au, body)
            elif typ == "vmlDrawing":
                vx = book.files[resolve(base, rel["Target"])].decode("utf-8", "replace")
                for shape in re.findall(r"<v:shape .*?</v:shape>", vx, re.S):
                    r = re.search(r"<x:Row>(\d+)</x:Row>", shape)
                    c = re.search(r"<x:Column>(\d+)</x:Column>", shape)
                    anc = re.search(r"<x:Anchor>\s*([^<]*)</x:Anchor>", shape)
                    if r and c and anc:
                        self.vml_shapes[(int(r.group(1)), int(c.group(1)))] = anc.group(1).strip()

    def col(self, header):
        return self.col_of.get(header)


# ───────────────────────── 產生 All ─────────────────────────

def build_rows(book, cat):
    name_col = cat.col(cat.name)
    task_col = cat.col("任務／地圖")
    if name_col is None or task_col is None:
        raise ValueError(f"「{cat.name}」找不到名稱欄或「任務／地圖」欄")
    ni, ti = col_to_num(name_col), col_to_num(task_col)
    skill_cols = [num_to_col(i) for i in range(ni + 1, ti)][:3]
    skill_cols += [None] * (3 - len(skill_cols))
    take_col = next((cat.col(h) for h in cat.names if h.startswith("接任地點")), None)
    pre_cols = [cat.col(h) for h in cat.names if h.startswith("前置發現")]

    rows = []
    for r in range(cat.first_row, cat.last_row + 1):
        name_cell = cat.cells.get((r, name_col))
        if name_cell is None or name_cell.value(book.shared) in (None, ""):
            continue
        src = {}   # All 欄名 → 來源欄字母（None = 空）
        src["★"] = cat.col("★")
        src["發現物"] = name_col
        for h, c in zip(["視", "考", "宗"], skill_cols):
            src[h] = c
        src["任務／地圖"] = task_col
        src["接任地點"] = take_col
        src["發現地點"] = cat.col("發現地點")
        for i, h in enumerate(["前置發現", "前置發現2", "前置發現3"]):
            src[h] = pre_cols[i] if i < len(pre_cols) else None
        src["備註"] = cat.col("備註")
        src["ID"] = cat.col("ID")
        src["版本"] = cat.col("版本")
        merged = None
        if len(pre_cols) > 3:
            vals = [cat.cells.get((r, c)) for c in pre_cols[2:]]
            vals = [v.value(book.shared) for v in vals if v is not None]
            vals = [v for v in vals if v not in (None, "")]
            if len(vals) > 1:
                merged = " / ".join(vals)
            elif len(vals) == 1:
                # 只有後面那欄有值 → 直接拿那一格
                for c in pre_cols[2:]:
                    cell = cat.cells.get((r, c))
                    if cell is not None and cell.value(book.shared) not in (None, ""):
                        src["前置發現3"] = c
                        break
        rows.append({"row": r, "src": src, "merged": merged,
                     "name": name_cell.value(book.shared),
                     "id": (cat.cells.get((r, src["ID"])) or Cell("", "")).value(book.shared)
                     if src["ID"] else None})
    return rows


class Styles:
    """在 styles.xml 的 cellXfs 加上「鎖定／不鎖定」版本，供 All 頁使用"""

    def __init__(self, book):
        self.book = book
        self.xml = book.text("xl/styles.xml")
        m = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", self.xml, re.S)
        self.xfs = re.findall(r"<xf\b[^>]*?/>|<xf\b[^>]*?>.*?</xf>", m.group(1), re.S)
        self.added = []
        self.cache = {}
        d = re.search(r"<dxfs[^>]*>(.*?)</dxfs>", self.xml, re.S)
        self.dxfs = re.findall(r"<dxf>.*?</dxf>|<dxf/>", d.group(1), re.S) if d else []
        self.has_dxfs_tag = bool(re.search(r"<dxfs\b", self.xml))
        self.dxf_added = []
        f = re.search(r"<fills[^>]*>(.*?)</fills>", self.xml, re.S)
        self.fills = re.findall(r"<fill>.*?</fill>|<fill/>", f.group(1), re.S) if f else []
        self.fill_added = []

    def dxf(self, xml):
        """加一個格式化條件用的格式，回傳 dxfId（重跑時沿用已有的）"""
        allx = self.dxfs + self.dxf_added
        if xml in allx:
            return allx.index(xml)
        self.dxf_added.append(xml)
        return len(allx)

    def _xf(self, idx):
        allx = self.xfs + self.added
        return allx[idx] if idx < len(allx) else allx[0]

    def filled(self, style, rgb):
        """回傳加了底色的格式編號"""
        idx = int(style) if style not in (None, "") else 0
        fill = (f'<fill><patternFill patternType="solid"><fgColor rgb="{rgb}"/>'
                '<bgColor indexed="64"/></patternFill></fill>')
        allf = self.fills + self.fill_added
        if fill not in allf:
            self.fill_added.append(fill)
            allf = self.fills + self.fill_added
        fid = allf.index(fill)
        xf = self._xf(idx)
        head = re.match(r"<xf\b[^>]*>", xf).group(0)
        new_head = re.sub(r'\s(fillId|applyFill)="\d+"', "", head)
        new_head = new_head.replace("<xf", f'<xf fillId="{fid}" applyFill="1"', 1)
        new = new_head + xf[len(head):]
        allx = self.xfs + self.added
        if new in allx:
            return str(allx.index(new))
        self.added.append(new)
        return str(len(allx))

    def _locked_state(self, xf):
        m = re.search(r'<protection[^>]*\blocked="(\d)"', xf)
        return m.group(1) != "0" if m else True      # 預設是鎖定

    def variant(self, style, locked):
        idx = int(style) if style not in (None, "") else 0
        key = (idx, locked)
        if key in self.cache:
            return self.cache[key]
        xf = self._xf(idx)
        if self._locked_state(xf) == locked:
            self.cache[key] = str(idx)
            return str(idx)
        hidden = re.search(r'<protection[^>]*\bhidden="(\d)"', xf)
        prot = f'<protection locked="{1 if locked else 0}"' + (
            f' hidden="{hidden.group(1)}"' if hidden else "") + "/>"
        new = re.sub(r"<protection[^>]*/>", "", xf)
        if new.endswith("/>"):
            new = new[:-2] + ">" + prot + "</xf>"
        else:
            new = new.replace("</xf>", prot + "</xf>")
        if "applyProtection=" in new:
            new = re.sub(r'applyProtection="\d"', 'applyProtection="1"', new, count=1)
        else:
            new = new.replace("<xf ", '<xf applyProtection="1" ', 1)
        # 重跑腳本時，上次加過的格式直接沿用，避免 styles.xml 越長越大
        all_xfs = self.xfs + self.added
        if new in all_xfs:
            self.cache[key] = str(all_xfs.index(new))
            return self.cache[key]
        new_idx = len(self.xfs) + len(self.added)
        self.added.append(new)
        self.cache[key] = str(new_idx)
        return str(new_idx)

    def save(self):
        xml = self.xml
        if self.added:
            total = len(self.xfs) + len(self.added)
            xml = re.sub(r'(<cellXfs[^>]*count=")\d+(")', rf"\g<1>{total}\2", xml, count=1)
            xml = xml.replace("</cellXfs>", "".join(self.added) + "</cellXfs>", 1)
        if self.fill_added:
            total = len(self.fills) + len(self.fill_added)
            xml = re.sub(r'(<fills[^>]*count=")\d+(")', rf"\g<1>{total}\2", xml, count=1)
            xml = xml.replace("</fills>", "".join(self.fill_added) + "</fills>", 1)
        if self.dxf_added:
            total = len(self.dxfs) + len(self.dxf_added)
            new = "".join(self.dxf_added)
            if re.search(r"<dxfs[^>]*/>", xml):
                xml = re.sub(r"<dxfs[^>]*/>", f'<dxfs count="{total}">{new}</dxfs>', xml, count=1)
            elif self.has_dxfs_tag:
                xml = re.sub(r'(<dxfs[^>]*count=")\d+(")', rf"\g<1>{total}\2", xml, count=1)
                xml = xml.replace("</dxfs>", new + "</dxfs>", 1)
            else:
                xml = xml.replace("</cellStyles>",
                                  f'</cellStyles><dxfs count="{total}">{new}</dxfs>', 1)
        if xml != self.xml:
            self.book.put("xl/styles.xml", xml)


def find_sheet(book, name):
    for i, s in enumerate(book.sheets):
        if s[0] == name:
            return i, s
    return None, None


def main():
    if len(sys.argv) > 1:
        src_path = sys.argv[1]
    else:
        print("請把 GVOexcel 檔拖到 bat 上，或：python gvo_add_all_sheet.py 檔名.xlsx")
        return 1
    if not os.path.exists(src_path):
        print(f"找不到檔案：{src_path}")
        return 1

    if src_path.lower().endswith(".bak.xlsx"):
        print("這是備份檔（.bak.xlsx），請改拖原本的 GVOexcel 檔。")
        return 1

    print(f"讀取：{src_path}")
    book = Book(src_path)
    names = [s[0] for s in book.sheets]
    if FIRST_SHEET not in names or LAST_SHEET not in names:
        print(f"找不到「{FIRST_SHEET}」或「{LAST_SHEET}」分頁")
        print(f"讀到的分頁（{len(names)} 個）：{names}")
        print("workbook.xml 開頭：", book.wb[:600])
        return 1
    cat_names = [n for n in names[names.index(FIRST_SHEET):names.index(LAST_SHEET) + 1]
                 if n != ALL_SHEET]
    cats = [Category(book, n, book.sheets[names.index(n)][2]) for n in cat_names]
    style_cat = next(c for c in cats if c.name == STYLE_SHEET)

    # 重建前記住舊 All 頁選的角色
    selected = 1
    _, old_all = find_sheet(book, ALL_SHEET)
    if old_all:
        oc = parse_cells(book.text(book.sheet_part(old_all[2]))).get(
            (split_ref(SELECT_CELL)[1], split_ref(SELECT_CELL)[0]))
        v = oc.value(book.shared) if oc else None
        if v is not None and str(v).strip() in ("1", "2", "3", "4", "5"):
            selected = int(str(v).strip())

    # ── 收集資料列 ──
    all_rows = []
    for cat in cats:
        rows = build_rows(book, cat)
        ids = [r["id"] for r in rows]
        for r in rows:
            r["cat"] = cat
            r["key_by_id"] = r["id"] not in (None, "") and ids.count(r["id"]) == 1
        print(f"  {cat.name:<6} {len(rows):>4} 筆")
        all_rows.extend(rows)
    print(f"  合計   {len(all_rows):>4} 筆")

    # ── 前置發現：建立對照表 ──
    def cell_text(cat, sr, col):
        if not col:
            return None
        c = cat.cells.get((sr, col))
        v = c.value(book.shared) if c is not None else None
        return None if v in (None, "") else str(v)

    # 預先算好的值跟 C1 選的角色一致（Excel 開檔會重算，但預覽／其他工具只看這個值）
    sel_col = CHECK_COLS[selected - 1]
    name_map, task_map = {}, {}       # 去空白的名稱 → [實際名稱, 所選角色是否完成]
    task_names = {}                   # 任務／地圖 → 它底下的發現物名稱
    name_rows = {}                    # 發現物名稱 → [(任務／地圖, 所選角色是否完成)]
    for rec in all_rows:
        cat, sr, src = rec["cat"], rec["row"], rec["src"]
        rec["task"] = cell_text(cat, sr, src["任務／地圖"])
        rec["tick1"] = bool((cell_text(cat, sr, cat.col(sel_col)) or "").strip())
        pres = [cell_text(cat, sr, src[h]) for h in PRE_COLS]
        if rec["merged"] is not None:
            pres[2] = rec["merged"]
        rec["pres"] = pres
        for mp, v in ((name_map, rec["name"]), (task_map, rec["task"])):
            if v is None:
                continue
            k = str(v).strip()
            e = mp.setdefault(k, [str(v), False])
            e[1] = e[1] or rec["tick1"]
        if rec["name"] is not None:
            name_rows.setdefault(str(rec["name"]).strip(), []).append(
                (rec["task"], rec["tick1"]))
        if rec["task"] is not None and rec["name"] is not None:
            task_names.setdefault(rec["task"].strip(), []).append(str(rec["name"]).strip())

    extras = []                       # (表格名, 欄名, {名稱: [實際名稱, 角色1是否完成]})
    for sname, keys in EXTRA_SHEETS:
        if sname not in names:
            continue
        try:
            ec = Category(book, sname, book.sheets[names.index(sname)][2])
        except (ValueError, KeyError):
            continue
        if any(ec.col(c) is None for c in CHECK_COLS):
            continue
        for k in keys:
            kc = ec.col(k)
            if kc is None:
                continue
            vals = {}
            for r in range(ec.first_row, ec.last_row + 1):
                v = cell_text(ec, r, kc)
                if v is None:
                    continue
                done = bool((cell_text(ec, r, ec.col(sel_col)) or "").strip())
                e = vals.setdefault(v.strip(), [v, False])
                e[1] = e[1] or done
            extras.append((ec.table, k, vals))

    def norm(t):
        """比對用：統一各種中點，去掉書名號、引號和空白"""
        t = re.sub(r"[‧·・˙•･]", "·", t)
        return re.sub(r"[《》〈〉「」『』\s]", "", t)

    def exact(c):
        if c in name_map:
            return (f"{ALL_TABLE}[發現物]", f"{ALL_TABLE}[[þ]:[5]]", *name_map[c])
        if c in task_map:
            return (f"{ALL_TABLE}[{sref_name('任務／地圖')}]", f"{ALL_TABLE}[[þ]:[5]]",
                    *task_map[c])
        for table, k, vals in extras:
            if c in vals:
                return (f"{table}[{sref_name(k)}]", f"{table}[[þ]:[5]]", *vals[c])
        return None

    norm_index = {}                   # 正規化後的名稱 → 原名稱（發現物優先）
    for mp in [name_map, task_map] + [v for _, _, v in extras]:
        for k in mp:
            norm_index.setdefault(norm(k), k)

    def near(a, b):
        """只差一個字、相鄰兩字對調、或多／少一個「的」「之」"""
        if len(a) == len(b):
            d = [i for i in range(len(a)) if a[i] != b[i]]
            if len(d) == 1:
                return True
            return (len(d) == 2 and d[1] == d[0] + 1
                    and a[d[0]] == b[d[1]] and a[d[1]] == b[d[0]])
        if abs(len(a) - len(b)) == 1:
            lo, sh = (a, b) if len(a) > len(b) else (b, a)
            return any(lo[i] in "的之" and lo[:i] + lo[i + 1:] == sh for i in range(len(lo)))
        return False

    fuzzy_log = {}                    # 前置欄原文 → (配到的名稱, 方式)

    def resolve_part(p):
        """回傳 None（不是前置）、("?",) 或 (比對範圍, 打勾範圍, 實際名稱, 所選角色是否完成)"""
        p = p.strip()
        if not p or any(w in p for w in SKIP_WORDS):
            return None
        q = re.split(r"[（(]", p)[0].strip()
        cands = [p] + ([q] if q and q != p else [])
        # 0. 同名的發現物不只一筆（例如渾天儀、犀牛）→ 用括號裡的地圖指定是哪一筆
        pm = re.search(r"[（(](.*)[)）]\s*$", p)
        if pm and len(name_rows.get(q, [])) > 1:
            for t in (pm.group(1).strip(), "圖：" + pm.group(1).strip()):
                hit = [r for r in name_rows[q] if r[0] is not None and r[0].strip() == t]
                if hit:
                    return (f"{ALL_TABLE}[發現物]", f"{ALL_TABLE}[[þ]:[5]]",
                            name_map[q][0], any(r[1] for r in hit),
                            f"{ALL_TABLE}[{sref_name('任務／地圖')}]", hit[0][0])
        for c in cands:                               # 1. 完全相同
            r = exact(c)
            if r:
                return r
        for c in cands:                               # 2. 只差標點
            k = norm_index.get(norm(c))
            if k:
                return exact(k)
        if q in NO_FUZZY or p in NO_FUZZY:
            return ("?",)
        if q in ALIAS_MAP and exact(ALIAS_MAP[q]):    # 3. 已知筆誤
            fuzzy_log[p] = (ALIAS_MAP[q], "筆誤表")
            return exact(ALIAS_MAP[q])
        m = re.search(r"[（(](.*)[)）]\s*$", p)       # 4. 括號裡的地圖 → 那張圖底下最像的發現物
        if m and q:
            for t in (m.group(1).strip(), "圖：" + m.group(1).strip()):
                group = task_names.get(t)
                if not group:
                    continue
                if len(set(group)) == 1:
                    best = [group[0]]
                else:
                    best = difflib.get_close_matches(q, sorted(set(group)), n=2, cutoff=0.6)
                    if len(best) == 2 and (difflib.SequenceMatcher(None, q, best[0]).ratio()
                                           == difflib.SequenceMatcher(None, q, best[1]).ratio()):
                        best = []
                if best and exact(best[0]):
                    fuzzy_log[p] = (best[0], "同一張地圖")
                    return exact(best[0])
        for c in cands:                               # 5. 只差一個字（要唯一）
            nc = norm(c)
            if len(nc) < 4:
                continue
            hits = [k for nk, k in norm_index.items() if near(nc, nk)]
            if len(hits) == 1:
                fuzzy_log[p] = (hits[0], "差一個字")
                return exact(hits[0])
        return ("?",)

    def criteria(text):
        t = text.replace("~", "~~").replace("*", "~*").replace("?", "~?").replace('"', '""')
        return f'"={t}"'

    stat = {"ok": 0, "?": 0}
    for rec in all_rows:
        helpers = []
        for i, text in enumerate(rec["pres"]):
            if text is None:
                helpers.append(None)
                continue
            parts = text.split(" / ") if (i == 2 and rec["merged"] is not None) else [text]
            res = [r for r in (resolve_part(x) for x in parts) if r is not None]
            if not res:
                helpers.append(None)
                continue
            known = [r for r in res if r[0] != "?"]
            unknown = len(known) < len(res)
            stat["ok"] += len(known)
            stat["?"] += len(res) - len(known)
            if not known:
                helpers.append(("const", "?"))
                continue
            terms = [f'COUNTIFS({r[0]},{criteria(r[2])},'
                     + (f'{r[4]},{criteria(r[5])},' if len(r) > 4 else "")
                     + f'INDEX({r[1]},0,{SELECT_CELL_ABS}),"?*")>0'
                     for r in known]
            cond = terms[0] if len(terms) == 1 else f"AND({','.join(terms)})"
            f = f'IF({cond},{chr(34) + "?" + chr(34) if unknown else 1},0)'
            done1 = all(r[3] for r in known)
            cached = ("?" if unknown else 1) if done1 else 0
            helpers.append(("formula", f, cached))
        rec["helpers"] = helpers
        c = [h[2] if h and h[0] == "formula" else (h[1] if h else None) for h in helpers]
        if rec["tick1"] or 0 in c:
            rec["can1"] = ""
        elif "?" in c:
            rec["can1"] = "?"
        else:
            rec["can1"] = "✔"
    print(f"  前置發現：自動判斷 {stat['ok']} 個，無法判斷 {stat['?']} 個")
    if fuzzy_log:
        print(f"  其中 {len(fuzzy_log)} 種寫法是近似比對（配錯的話，把左邊的名稱加進腳本開頭的 NO_FUZZY）：")
        for k, (v, how) in sorted(fuzzy_log.items()):
            print(f"    {k}  →  {v}（{how}）")

    # ── 預設格式（史跡頁）──
    hdr_style, def_style, width = {}, {}, {}
    cols_xml = re.search(r"<cols>(.*?)</cols>", style_cat.xml, re.S)
    col_widths = {}
    if cols_xml:
        for tag in re.findall(r"<col [^>]*?/>", cols_xml.group(1)):
            a = attrs(tag)
            for n in range(int(a["min"]), int(a["max"]) + 1):
                col_widths[num_to_col(n)] = a.get("width")
    for h in OUT_HEADERS:
        sc = style_cat.col(STYLE_MAP[h])
        hc = style_cat.cells.get((style_cat.header_row, sc)) if sc else None
        dc = style_cat.cells.get((style_cat.first_row, sc)) if sc else None
        hdr_style[h] = hc.style if hc else None
        def_style[h] = dc.style if dc else None
        width[h] = EXTRA_WIDTH.get(h) or (col_widths.get(sc) if sc else None) or 10
    ver_hdr_style = hdr_style["版本"]
    hdr_style["分類"] = ver_hdr_style

    out_col = {h: num_to_col(i + 1) for i, h in enumerate(OUT_HEADERS)}
    last_col = out_col[OUT_HEADERS[-1]]
    n_rows = len(all_rows) + HDR      # 表格最後一列

    # ── 工作表 rels（超連結、註解、表格）──
    sheet_rels = []          # (rId, type, target, external)

    rel_cache = {}

    def add_rel(typ, target, external=False):
        """target 為未轉義的原始字串；同一網址只建一個關聯"""
        key = (typ, target, external)
        if key in rel_cache:
            return rel_cache[key]
        rid = f"rId{len(sheet_rels) + 1}"
        sheet_rels.append((rid, typ, target, external))
        rel_cache[key] = rid
        return rid

    # ── 組 sheetData ──
    styles = Styles(book)
    locked_cols = {out_col[h] for h in CHECK_COLS + ["可接"] + HELPER_COLS}

    def c_xml(ref, style, attr_extra="", inner=""):
        col, row = split_ref(ref)
        locked = (row <= HDR and ref != SELECT_CELL) or col in locked_cols
        style = styles.variant(style, locked=locked)
        s = f' s="{style}"' if style != "0" else ""
        if inner:
            return f'<c r="{ref}"{s}{attr_extra}>{inner}</c>'
        return f'<c r="{ref}"{s}{attr_extra}/>'

    def inline(ref, style, text):
        return c_xml(ref, style, ' t="inlineStr"',
                     f'<is><t xml:space="preserve">{escape(str(text))}</t></is>')

    out = []
    hyperlinks = []
    comments = []            # (ref, author, body, anchor_delta, row0, col0)
    note = ("← 選 1～5（1＝þ 欄）　前置發現：綠＝已完成、紅＝未完成　"
            "可接：✔＝前置都完成、?＝有前置無法判斷，請自行確認")
    note = "← 黃色格選 1～5" + note[len("← 選 1～5"):]
    out.append(f'<row r="1">'
               f'{inline("B1", def_style["發現物"], "查看角色 →")}'
               f'{c_xml(SELECT_CELL, styles.filled(def_style["★"], "FFFFE699"), "", f"<v>{selected}</v>")}'
               f'{inline("D1", def_style["發現物"], note)}</row>')
    row_parts = [inline(f"{out_col[h]}{HDR}", hdr_style[h], h) for h in OUT_HEADERS]
    out.append(f'<row r="{HDR}">{"".join(row_parts)}</row>')

    # 視／考／宗 三欄在各分類放的技能不同，在表頭加一則註解說明
    groups = {}
    for cat in cats:
        ni = col_to_num(cat.col(cat.name)); ti = col_to_num(cat.col("任務／地圖"))
        key = "／".join(n for n in cat.names if ni < col_to_num(cat.col(n)) < ti)
        groups.setdefault(key, []).append(cat.name)
    tip = "這三欄（視、考、宗）在各分類放的是：\n" + "\n".join(
        f"{k}：{'、'.join(v)}" for k, v in groups.items())
    hc = col_to_num(out_col["視"]) - 1
    comments.append((f'{out_col["視"]}{HDR}', "",
                     f'<text><t xml:space="preserve">{escape(tip)}</t></text>',
                     f"{hc + 1}, 15, {HDR - 1}, 10, {hc + 8}, 15, {HDR + 4 + len(groups)}, 4",
                     HDR - 1, hc, HDR - 1, hc))

    for i, rec in enumerate(all_rows, start=HDR + 1):
        cat, sr, src = rec["cat"], rec["row"], rec["src"]
        parts = []
        for h in OUT_HEADERS:
            ref = f"{out_col[h]}{i}"
            if h == "分類":
                vs = src["版本"]
                vc = cat.cells.get((sr, vs)) if vs else None
                parts.append(inline(ref, vc.style if vc and vc.style else def_style[h], cat.name))
                continue
            if h in CHECK_COLS:
                sc = cat.col(h)
                cell = cat.cells.get((sr, sc)) if sc else None
                style = cell.style if cell and cell.style else def_style[h]
                if sc is None:
                    parts.append(c_xml(ref, style))
                    continue
                if rec["key_by_id"]:
                    key_here, key_there = "ID", "ID"
                else:
                    key_here, key_there = "發現物", cat.name
                tgt = f"{cat.table}[{sref_name(h)}]"
                match = (f"MATCH({ALL_TABLE}[[#This Row],[{sref_name(key_here)}]],"
                         f"{cat.table}[{sref_name(key_there)}],0)")
                f = f'IFERROR(IF(INDEX({tgt},{match})="","",INDEX({tgt},{match})),"")'
                cached = cell.value(book.shared) if cell else None
                cached = "" if cached is None else cached
                parts.append(c_xml(ref, style, ' t="str"',
                                   f"<f>{escape(f)}</f><v>{escape(str(cached))}</v>"))
                continue
            if h == "可接":
                tr = f"{ALL_TABLE}[[#This Row],[þ]:[5]]"
                hr = f"{ALL_TABLE}[[#This Row],[{HELPER_COLS[0]}]:[{HELPER_COLS[-1]}]]"
                f = (f'IF(INDEX({tr},1,{SELECT_CELL_ABS})<>"","",'
                     f'IF(COUNTIF({hr},0)>0,"",IF(COUNTIF({hr},"~?")>0,"?","✔")))')
                parts.append(c_xml(ref, def_style[h], ' t="str"',
                                   f"<f>{escape(f)}</f><v>{escape(rec['can1'])}</v>"))
                continue
            if h in HELPER_COLS:
                hp = rec["helpers"][HELPER_COLS.index(h)]
                if hp is None:
                    parts.append(c_xml(ref, def_style[h]))
                elif hp[0] == "const":
                    parts.append(inline(ref, def_style[h], hp[1]))
                elif hp[2] == "?":
                    parts.append(c_xml(ref, def_style[h], ' t="str"',
                                       f"<f>{escape(hp[1])}</f><v>?</v>"))
                else:
                    parts.append(c_xml(ref, def_style[h], "",
                                       f"<f>{escape(hp[1])}</f><v>{hp[2]}</v>"))
                continue
            sc = src.get(h)
            cell = cat.cells.get((sr, sc)) if sc else None
            # 前置發現3、4 併成一格時，內容用合併後的文字，超連結和註解照樣帶過來
            merged_here = h == "前置發現3" and rec["merged"] is not None
            if merged_here:
                parts.append(inline(ref, cell.style if cell is not None and cell.style
                                    else def_style[h], rec["merged"]))
            if cell is None:
                if not merged_here:
                    parts.append(c_xml(ref, def_style[h]))
                continue
            style = cell.style or def_style[h]
            at = re.sub(r'\s*\bs="\d+"', "", cell.attrs)
            inner = cell.inner
            if merged_here:
                pass
            elif cell.has_formula():
                # 來源是公式 → 只留結果
                v = cell.value(book.shared)
                if v in (None, ""):
                    parts.append(c_xml(ref, style))
                else:
                    parts.append(inline(ref, style, v))
            else:
                parts.append(c_xml(ref, style, at, inner))

            # 超連結
            src_ref = f"{sc}{sr}"
            link = cat.links.get(src_ref)
            if link:
                extra = ""
                if "location" in link:
                    extra += f' location="{link["location"]}"'
                if "display" in link:
                    extra += f' display="{link["display"]}"'
                if "tooltip" in link:
                    extra += f' tooltip="{link["tooltip"]}"'
                rid_attr = ""
                if "r:id" in link and link["r:id"] in cat.rels:
                    rel = cat.rels[link["r:id"]]
                    rid = add_rel("hyperlink", unescape_attr(rel["Target"]),
                                  rel.get("TargetMode") == "External")
                    rid_attr = f' r:id="{rid}"'
                if rid_attr or extra:
                    hyperlinks.append(f'<hyperlink ref="{ref}"{rid_attr}{extra}/>')

            # 註解
            if src_ref in cat.comments:
                au, body = cat.comments[src_ref]
                r0, c0 = sr - 1, col_to_num(sc) - 1
                anchor = cat.vml_shapes.get((r0, c0))
                comments.append((ref, au, body, anchor, r0, c0, i - 1,
                                 col_to_num(out_col[h]) - 1))
        out.append(f'<row r="{i}">{"".join(parts)}</row>')

    # ── 格式化條件（接任地點的顏色）──
    cf_parts = []
    for block in re.findall(r"<conditionalFormatting .*?</conditionalFormatting>",
                            style_cat.xml, re.S):
        sq = attrs(re.match(r"<conditionalFormatting [^>]*>", block).group(0))["sqref"]
        m = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", sq)
        if not m or m.group(1) != m.group(3):
            continue
        src_letter = m.group(1)
        hdr_name = next((n for n in style_cat.names if style_cat.col(n) == src_letter), None)
        target_h = next((h for h in OUT_HEADERS if STYLE_MAP[h] == hdr_name
                         and h not in ("分類",)), None)
        if not target_h:
            continue
        new_letter = out_col[target_h]
        # 套用範圍改成 All 頁的資料列，公式裡的列號跟著平移
        delta = (HDR + 1) - int(m.group(2))
        block = block.replace(f'sqref="{sq}"',
                              f'sqref="{new_letter}{HDR + 1}:{new_letter}{n_rows}"')
        block = re.sub(rf"\b{src_letter}(\d+)\b",
                       lambda mm: f"{new_letter}{int(mm.group(1)) + delta}", block)
        cf_parts.append(block)

    # 其他分類頁的「接任地點」如果有史跡頁沒有的規則（例如海洋生物的「釣魚」），也補進來
    take_letter = out_col["接任地點"]
    seen_text = set(re.findall(r'<cfRule [^>]*\btext="([^"]*)"', "".join(cf_parts)))
    extra_rules = []
    for cat in cats:
        if cat is style_cat:
            continue
        src_letter = next((cat.col(n) for n in cat.names if n.startswith("接任地點")), None)
        if not src_letter:
            continue
        for block in re.findall(r"<conditionalFormatting .*?</conditionalFormatting>",
                                cat.xml, re.S):
            sq = attrs(re.match(r"<conditionalFormatting [^>]*>", block).group(0))["sqref"]
            m = re.match(r"([A-Z]+)(\d+)", sq)
            if not m or m.group(1) != src_letter:
                continue
            delta = (HDR + 1) - int(m.group(2))
            for rule in re.findall(r"<cfRule [^>]*?(?:/>|>.*?</cfRule>)", block, re.S):
                text = attrs(re.match(r"<cfRule [^>]*>", rule).group(0)).get("text")
                if text is None or text in seen_text:
                    continue
                seen_text.add(text)
                rule = re.sub(rf"\b{src_letter}(\d+)\b",
                              lambda mm: f"{take_letter}{int(mm.group(1)) + delta}", rule)
                rule = re.sub(r'priority="\d+"', f'priority="{50 + len(extra_rules)}"', rule)
                extra_rules.append(rule)
    if extra_rules:
        cf_parts.append(f'<conditionalFormatting sqref="{take_letter}{HDR + 1}:'
                        f'{take_letter}{n_rows}">{"".join(extra_rules)}</conditionalFormatting>')

    # 前置發現上色：綠＝已完成、紅＝未完成（依隱藏的 前1～前3 欄）
    green = styles.dxf('<dxf><font><color rgb="FF006100"/></font><fill><patternFill>'
                       '<bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>')
    red = styles.dxf('<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill>'
                     '<bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>')
    prio = 100
    first = HDR + 1
    for pc, hc in zip(PRE_COLS, HELPER_COLS):
        L, H = out_col[pc], out_col[hc]
        cf_parts.append(
            f'<conditionalFormatting sqref="{L}{first}:{L}{n_rows}">'
            f'<cfRule type="expression" dxfId="{green}" priority="{prio}">'
            f'<formula>AND(ISNUMBER(${H}{first}),${H}{first}=1)</formula></cfRule>'
            f'<cfRule type="expression" dxfId="{red}" priority="{prio + 1}">'
            f'<formula>AND(ISNUMBER(${H}{first}),${H}{first}=0)</formula></cfRule>'
            '</conditionalFormatting>')
        prio += 2
    validation = ('<dataValidations count="1"><dataValidation type="list" allowBlank="0" '
                  f'showErrorMessage="1" errorTitle="查看角色" error="請選 1～5" '
                  f'sqref="{SELECT_CELL}"><formula1>"1,2,3,4,5"</formula1>'
                  '</dataValidation></dataValidations>')

    # ── 檔名編號 ──
    def next_num(prefix, ext):
        nums = [int(m.group(1)) for n in book.files
                for m in [re.fullmatch(rf"{re.escape(prefix)}(\d+){re.escape(ext)}", n)] if m]
        return max(nums, default=0) + 1

    # ── 移除舊的 All 頁 ──
    old_idx, old = find_sheet(book, ALL_SHEET)
    ct = book.text("[Content_Types].xml")
    wb = book.wb
    wb_rels_text = book.text("xl/_rels/workbook.xml.rels")
    if old:
        part = book.sheet_part(old[2])
        rp = rels_path(part)
        if rp in book.files:
            base = part.rsplit("/", 1)[0]
            for rel in parse_rels(book.text(rp)).values():
                if rel.get("TargetMode") == "External":
                    continue
                t = resolve(base, rel["Target"])
                book.remove(t)
                book.remove(rels_path(t))
                ct = re.sub(rf'<Override PartName="/{re.escape(t)}"[^>]*/>', "", ct)
            book.remove(rp)
        book.remove(part)
        ct = re.sub(rf'<Override PartName="/{re.escape(part)}"[^>]*/>', "", ct)
        wb = wb.replace(old[3], "")
        wb_rels_text = re.sub(rf'<Relationship Id="{old[2]}"[^>]*/>', "", wb_rels_text)

        def fix_local(m):
            n = int(m.group(2))
            return m.group(0) if n < old_idx else f'{m.group(1)}{n - 1}"'
        wb = re.sub(rf'<definedName [^>]*localSheetId="{old_idx}"[^>]*>.*?</definedName>',
                    "", wb, flags=re.S)
        wb = re.sub(r'(localSheetId=")(\d+)"', fix_local, wb)
        wb = re.sub(r"<definedNames>\s*</definedNames>", "", wb)

    # ── calcChain 刪掉讓 Excel 重建，並要求開檔時重算 ──
    for rid, rel in parse_rels(wb_rels_text).items():
        if rel["Type"].endswith("/calcChain"):
            book.remove(resolve("xl", rel["Target"]))
            wb_rels_text = re.sub(rf'<Relationship Id="{rid}"[^>]*/>', "", wb_rels_text)
            ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", ct)
    if "fullCalcOnLoad" not in wb:
        wb = re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', wb, count=1)

    # ── 新檔案 ──
    sheet_no = next_num("xl/worksheets/sheet", ".xml")
    sheet_part = f"xl/worksheets/sheet{sheet_no}.xml"
    table_no = next_num("xl/tables/table", ".xml")
    table_part = f"xl/tables/table{table_no}.xml"
    table_id = max([int(attrs(re.search(r"<table [^>]*>", book.text(n)).group(0))["id"])
                    for n in book.files if re.fullmatch(r"xl/tables/table\d+\.xml", n)],
                   default=0) + 1

    # 表格 XML
    style_info = re.search(r"<tableStyleInfo [^>]*?/>", style_cat.table_xml)
    style_info = style_info.group(0) if style_info else (
        '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" '
        'showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>')
    tcols = "".join(f'<tableColumn id="{i + 1}" name="{escape(h, {chr(34): "&quot;"})}"/>'
                    for i, h in enumerate(OUT_HEADERS))
    table_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<table xmlns="{NS_MAIN}" id="{table_id}" name="{ALL_TABLE}" '
        f'displayName="{ALL_TABLE}" ref="A{HDR}:{last_col}{n_rows}">'
        f'<autoFilter ref="A{HDR}:{last_col}{n_rows}"/>'
        f'<tableColumns count="{len(OUT_HEADERS)}">{tcols}</tableColumns>'
        f'{style_info}</table>')
    book.put(table_part, table_xml)
    table_rid = add_rel("table", f"../tables/table{table_no}.xml")

    # 註解
    legacy = ""
    if comments:
        cm_no = next_num("xl/comments", ".xml")
        vml_no = next_num("xl/drawings/vmlDrawing", ".vml")
        authors = []
        for c in comments:
            if c[1] not in authors:
                authors.append(c[1])
        clist = "".join(
            f'<comment ref="{c[0]}" authorId="{authors.index(c[1])}">{c[2]}</comment>'
            for c in comments)
        book.put(f"xl/comments{cm_no}.xml",
                 '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 f'<comments xmlns="{NS_MAIN}"><authors>'
                 + "".join(f"<author>{escape(a)}</author>" for a in authors)
                 + f"</authors><commentList>{clist}</commentList></comments>")
        idmaps = []
        for n in book.files:
            if n.endswith(".vml"):
                for m in re.findall(r'idmap v:ext="edit" data="([\d,]+)"',
                                    book.files[n].decode("utf-8", "replace")):
                    idmaps += [int(x) for x in m.split(",")]
        need = (len(comments) // 1024) + 1
        first_map = max(idmaps, default=0) + 1
        maps = list(range(first_map, first_map + need))
        shapes = []
        for k, (ref, au, body, anchor, r0, c0, nr0, nc0) in enumerate(comments):
            if anchor:
                a = [int(x) for x in re.split(r"\s*,\s*", anchor)]
                dc, dr = a[0] - c0, a[2] - r0
                a = [nc0 + dc, a[1], nr0 + dr, a[3],
                     nc0 + dc + (a[4] - a[0]), a[5], nr0 + dr + (a[6] - a[2]), a[7]]
                a = [max(0, x) for x in a]
            else:
                a = [nc0 + 1, 15, max(0, nr0 - 1), 10, nc0 + 4, 15, nr0 + 5, 4]
            sid = maps[0] * 1024 + k + 1
            shapes.append(
                f'<v:shape id="_x0000_s{sid}" type="#_x0000_t202" '
                'style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;'
                'width:216pt;height:120pt;z-index:1;visibility:hidden" '
                'fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/>'
                '<v:shadow on="t" color="black" obscured="t"/>'
                '<v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto">'
                '<div style="text-align:left"></div></v:textbox>'
                '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
                f'<x:Anchor>{", ".join(map(str, a))}</x:Anchor>'
                f'<x:AutoFill>False</x:AutoFill><x:Row>{nr0}</x:Row>'
                f'<x:Column>{nc0}</x:Column></x:ClientData></v:shape>')
        vml = ('<xml xmlns:v="urn:schemas-microsoft-com:vml" '
               'xmlns:o="urn:schemas-microsoft-com:office:office" '
               'xmlns:x="urn:schemas-microsoft-com:office:excel">'
               f'<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" '
               f'data="{",".join(map(str, maps))}"/></o:shapelayout>'
               '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" '
               'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
               '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
               + "".join(shapes) + "</xml>")
        book.put(f"xl/drawings/vmlDrawing{vml_no}.vml", vml)
        add_rel("comments", f"../comments{cm_no}.xml")
        vml_rid = add_rel("vmlDrawing", f"../drawings/vmlDrawing{vml_no}.vml")
        legacy = f'<legacyDrawing r:id="{vml_rid}"/>'
        ct = ct.replace("</Types>",
                        f'<Override PartName="/xl/comments{cm_no}.xml" '
                        f'ContentType="{CT_COMMENTS}"/></Types>')
        if 'Extension="vml"' not in ct:
            ct = ct.replace("</Types>",
                            '<Default Extension="vml" ContentType="application/'
                            'vnd.openxmlformats-officedocument.vmlDrawing"/></Types>')

    # 工作表 XML
    fmt = re.search(r"<sheetFormatPr [^>]*?/>", style_cat.xml)
    fmt = fmt.group(0) if fmt else '<sheetFormatPr defaultRowHeight="16.5"/>'
    fmt = re.sub(r'\s*x14ac:\w+="[^"]*"', "", fmt)
    cols = "".join(
        f'<col min="{i + 1}" max="{i + 1}" width="{width[h]}" '
        f'style="{styles.variant(0, locked=out_col[h] in locked_cols)}" customWidth="1"'
        + (' hidden="1"' if h in HELPER_COLS else "") + "/>"
        for i, h in enumerate(OUT_HEADERS))
    # 表格右邊的空白欄也保持可編輯
    cols += (f'<col min="{len(OUT_HEADERS) + 1}" max="16384" width="9" '
             f'style="{styles.variant(0, locked=False)}"/>')
    # 保護：允許選取、調整格式、欄寬列高、篩選；不設密碼
    protection = ('<sheetProtection sheet="1" objects="1" scenarios="1" '
                  'formatCells="0" formatColumns="0" formatRows="0" autoFilter="0"/>')
    freeze = out_col["þ"]   # 凍結：表頭列 + 到「任務／地圖」為止的欄
    xsplit = col_to_num(freeze) - 1
    tab = re.search(r"<tabColor [^>]*?/>", style_cat.xml)
    sheet_pr = f"<sheetPr>{tab.group(0)}</sheetPr>" if tab else ""
    links_xml = f"<hyperlinks>{''.join(hyperlinks)}</hyperlinks>" if hyperlinks else ""
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_R}">{sheet_pr}'
        f'<dimension ref="A1:{last_col}{n_rows}"/>'
        '<sheetViews><sheetView tabSelected="1" workbookViewId="0">'
        f'<pane xSplit="{xsplit}" ySplit="{HDR}" topLeftCell="{freeze}{HDR + 1}" '
        'activePane="bottomRight" state="frozen"/>'
        f'<selection pane="topRight" activeCell="{freeze}1" sqref="{freeze}1"/>'
        f'<selection pane="bottomLeft" activeCell="A{HDR + 1}" sqref="A{HDR + 1}"/>'
        f'<selection pane="bottomRight" activeCell="{freeze}{HDR + 1}" '
        f'sqref="{freeze}{HDR + 1}"/>'
        '</sheetView></sheetViews>'
        f'{fmt}<cols>{cols}</cols><sheetData>{"".join(out)}</sheetData>'
        f'{protection}{"".join(cf_parts)}{validation}{links_xml}'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" '
        'header="0.3" footer="0.3"/>'
        f'{legacy}<tableParts count="1"><tablePart r:id="{table_rid}"/></tableParts>'
        '</worksheet>')
    book.put(sheet_part, sheet_xml)
    styles.save()

    rel_xml = "".join(
        f'<Relationship Id="{rid}" Type="{REL_BASE}{typ}" '
        f'Target="{escape(tgt, {chr(34): "&quot;"})}"'
        + (' TargetMode="External"' if ext else "") + "/>"
        for rid, typ, tgt, ext in sheet_rels)
    book.put(rels_path(sheet_part),
             '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
             f'relationships">{rel_xml}</Relationships>')

    # Content types
    ct = ct.replace("</Types>",
                    f'<Override PartName="/{sheet_part}" ContentType="{CT_SHEET}"/>'
                    f'<Override PartName="/{table_part}" ContentType="{CT_TABLE}"/></Types>')
    book.put("[Content_Types].xml", ct)

    # workbook：放到第一頁
    wb_rid_nums = [int(x) for x in re.findall(r'Id="rId(\d+)"', wb_rels_text)]
    new_rid = f"rId{max(wb_rid_nums, default=0) + 1}"
    wb_rels_text = wb_rels_text.replace(
        "</Relationships>",
        f'<Relationship Id="{new_rid}" Type="{REL_BASE}worksheet" '
        f'Target="worksheets/sheet{sheet_no}.xml"/></Relationships>')
    book.put("xl/_rels/workbook.xml.rels", wb_rels_text)

    sheet_ids = [int(x) for x in re.findall(r'<sheet [^>]*sheetId="(\d+)"', wb)]
    new_sheet = f'<sheet name="{ALL_SHEET}" sheetId="{max(sheet_ids) + 1}" r:id="{new_rid}"/>'
    wb = wb.replace("<sheets>", "<sheets>" + new_sheet, 1)
    wb = re.sub(r'(localSheetId=")(\d+)"', lambda m: f'{m.group(1)}{int(m.group(2)) + 1}"', wb)
    wb = re.sub(r'\s(activeTab|firstSheet)="\d+"', "", wb)   # 開檔停在 All 頁
    book.put("xl/workbook.xml", wb)

    # 其他頁取消「選取中」，避免同時選到多頁
    for s in book.sheets:
        if s[0] == ALL_SHEET:
            continue
        part = book.sheet_part(s[2])
        if part in book.files:
            x = book.text(part)
            if 'tabSelected="1"' in x:
                book.put(part, x.replace(' tabSelected="1"', ""))

    # ── 存檔 ──
    bak = os.path.splitext(src_path)[0] + ".bak.xlsx"
    try:
        shutil.copy2(src_path, bak)
        book.save(src_path)
    except PermissionError:
        print("\n無法寫入，檔案可能正在 Excel 裡開著，請先關閉再試。")
        return 1
    print(f"\n完成：已在原檔加入「{ALL_SHEET}」頁（{len(all_rows)} 筆、"
          f"{len(hyperlinks)} 個超連結、{len(comments)} 則註解）")
    print(f"備份：{bak}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

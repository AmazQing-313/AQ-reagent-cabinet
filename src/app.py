# -*- coding: utf-8 -*-
"""
试剂柜速查 · 桌面版
- 纯本地窗口程序，不需要浏览器、不联网。
- 数据保存在 exe 同目录的 data.json。
"""
import csv
import json
import os
import re
import shutil
import sys
import uuid
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_NAME = "试剂柜速查"
FIELDS = ("name", "alias", "cas", "spec", "category", "cabinet", "slot", "qty", "note")
LABELS = {
    "name": "名称", "alias": "别名 / 化学式", "cas": "CAS 号", "spec": "规格",
    "category": "类型", "cabinet": "柜号", "slot": "柜内位置", "qty": "数量", "note": "备注",
}
CABINETS = ["一号柜", "二号柜", "三号柜", "四号柜"]
CATS = ["无机盐", "有机溶剂", "酸碱", "指示剂", "氧化剂", "还原剂", "生物试剂", "其他"]

# 导入时允许的表头别名（写 Excel 时不用刻意对齐我们的列名）
HEADER_ALIASES = {
    "name": ("名称", "品名", "试剂名称", "材料名称", "物料名称", "中文名", "试剂", "材料", "name"),
    "alias": ("别名", "英文名", "化学式", "分子式", "英文", "又名", "alias"),
    "cas": ("cas号", "cas", "casno", "cas编号", "cas no"),
    "spec": ("规格", "型号", "包装", "纯度", "规格型号", "spec"),
    "category": ("类型", "分类", "类别", "种类", "category"),
    "cabinet": ("柜号", "柜子", "柜", "存放柜", "柜位"),
    "slot": ("柜内位置", "位置", "存放位置", "层", "层位", "具体位置", "slot"),
    "qty": ("数量", "库存", "剩余", "余量", "qty"),
    "note": ("备注", "说明", "note", "remark"),
}


def normalize_header(text):
    """去掉括号备注、空格和结尾标点，统一小写，便于和别名表匹配。"""
    s = str(text or "")
    s = re.sub(r"[（(].*?[)）]", "", s)
    s = s.replace(" ", "").replace("\u3000", "").replace("\t", "")
    return s.strip(" .。:*：、,，;；").lower()


ALIAS_INDEX = {}
for _key, _names in HEADER_ALIASES.items():
    for _n in _names:
        ALIAS_INDEX.setdefault(normalize_header(_n), _key)


def map_columns(headers):
    mapping = {}
    for idx, h in enumerate(headers):
        key = ALIAS_INDEX.get(normalize_header(h))
        if key and key not in mapping:
            mapping[key] = idx
    return mapping


def detect_header(rows, max_scan=10):
    """在前 10 行里找最像表头的一行（必须有「名称」）。返回 (行号, 列映射)。"""
    best = None
    for i, row in enumerate(rows[:max_scan]):
        m = map_columns(row)
        if "name" not in m:
            continue
        score = len(m)
        if best is None or score > best[1]:
            best = (i, score, m)
    if not best:
        return None
    return best[0], best[2]


def split_location(cabinet, slot):
    """「一号柜 二层 左」这种合并写法，拆成柜号 + 柜内位置。"""
    cabinet = (cabinet or "").strip()
    slot = (slot or "").strip()
    if not cabinet and "柜" in slot:
        i = slot.index("柜")
        return slot[: i + 1].strip(), slot[i + 1:].strip(" ·-—,，、/")
    return cabinet, slot


def _col_index(ref):
    """A1 → 0, C3 → 2。"""
    letters = "".join(ch for ch in str(ref) if ch.isalpha())
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def _sheet_rows(xml_bytes, shared):
    """解析一张 worksheet 的 XML，返回二维字符串列表。"""
    import xml.etree.ElementTree as ET

    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    root = ET.fromstring(xml_bytes)
    data = root.find(NS + "sheetData")
    rows = []
    if data is None:
        return rows
    for row in data.findall(NS + "row"):
        cells = {}
        for c in row.findall(NS + "c"):
            idx = _col_index(c.get("r") or "")
            if idx < 0:
                continue
            t = c.get("t")
            v = c.find(NS + "v")
            if t == "s" and v is not None and v.text is not None:
                try:
                    val = shared[int(v.text)]
                except (ValueError, IndexError):
                    val = ""
            elif t == "inlineStr":
                val = "".join(x.text or "" for x in c.iter(NS + "t"))
            elif t == "b" and v is not None:
                val = "TRUE" if v.text == "1" else "FALSE"
            else:
                val = v.text if v is not None and v.text is not None else ""
            cells[idx] = str(val).strip()
        if cells:
            width = max(cells) + 1
            rows.append([cells.get(i, "") for i in range(width)])
        else:
            rows.append([])
    return rows


def read_xlsx(path):
    """用标准库直接读 .xlsx（不依赖 openpyxl，避免把 numpy/Pillow 一起打包）。"""
    import xml.etree.ElementTree as ET
    import zipfile

    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"

    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall(NS + "si"):
                shared.append("".join(t.text or "" for t in si.iter(NS + "t")))

        rels = {}
        if "xl/_rels/workbook.xml.rels" in names:
            rr = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            for rel in rr.findall(PKG + "Relationship"):
                rels[rel.get("Id")] = rel.get("Target", "")

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheets_node = wb.find(NS + "sheets")
        if sheets_node is None:
            return []

        candidates = []
        for sh in sheets_node.findall(NS + "sheet"):
            target = rels.get(sh.get(RID), "")
            if not target:
                continue
            target = target.lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            if target in names:
                candidates.append(target)

        sheets = []
        for target in candidates[:8]:
            rows = _sheet_rows(z.read(target), shared)
            if any(any(c for c in r) for r in rows):
                sheets.append(rows)
        if not sheets:
            return []
        for rows in sheets:  # 优先用有表头的那张表
            if detect_header(rows):
                return rows
        return sheets[0]


def read_table(path):
    """读 .xlsx/.xlsm 或 .csv，返回二维字符串列表。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return read_xlsx(path)
    if ext == ".xls":
        raise RuntimeError("不支持旧版 .xls。请在 Excel 里「另存为」.xlsx 或 CSV 后再导入。")
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        return [list(r) for r in csv.reader(f)]


def rows_to_items(rows, header_row, mapping):
    items = []
    for row in rows[header_row + 1:]:
        def cell(key):
            i = mapping.get(key)
            if i is None or i >= len(row) or row[i] is None:
                return ""
            return str(row[i]).strip()

        name = cell("name")
        if not name:
            continue
        cabinet, slot = split_location(cell("cabinet"), cell("slot"))
        items.append({
            "name": name, "alias": cell("alias"), "cas": cell("cas"), "spec": cell("spec"),
            "category": cell("category"), "cabinet": cabinet, "slot": slot,
            "qty": cell("qty"), "note": cell("note"),
        })
    return items



def new_id():
    """条目 id：同一毫秒批量导入也不会重号。"""
    return "i" + uuid.uuid4().hex[:12]


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


BASE = app_dir()
DATA_FILE = os.path.join(BASE, "data.json")
BACKUP_DIR = os.path.join(BASE, "备份")


def now_text():
    return datetime.now().strftime("%H:%M:%S")


def load_items():
    if not os.path.exists(DATA_FILE):
        return []
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            raw = json.load(f)
        items = raw.get("items") if isinstance(raw, dict) else raw
        if not isinstance(items, list):
            raise ValueError("items 不是数组")
        out = []
        for it in items:
            if not isinstance(it, dict) or not str(it.get("name", "")).strip():
                continue
            row = {k: str(it.get(k) if it.get(k) is not None else "").strip() for k in FIELDS}
            row["id"] = str(it.get("id") or new_id())
            row["updatedAt"] = int(it.get("updatedAt") or 0)
            out.append(row)
        return out
    except Exception as exc:
        broken = DATA_FILE + ".broken-" + datetime.now().strftime("%Y%m%d-%H%M%S")
        try:
            shutil.copy2(DATA_FILE, broken)
        except Exception:
            pass
        print("读取 data.json 失败：", repr(exc))
        return []


def write_items(items):
    tmp = DATA_FILE + ".tmp"
    payload = {"app": "reagent-locator", "savedAt": datetime.now().isoformat(timespec="seconds"), "items": items}
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, DATA_FILE)


def daily_backup():
    if not os.path.exists(DATA_FILE):
        return
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        target = os.path.join(BACKUP_DIR, "data-" + datetime.now().strftime("%Y%m%d") + ".json")
        if not os.path.exists(target):
            shutil.copy2(DATA_FILE, target)
    except Exception:
        pass


def cab_index(cab):
    s = str(cab or "")
    digits = "".join(ch for ch in s if ch.isdigit())
    if digits:
        return int(digits)
    for i, ch in enumerate("零一二三四五六七八九十"):
        if i and ch in s:
            return i
    return 99


CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def slot_key(slot):
    """把「二层 左」这类位置转成可排序的数字序列，避免中文按笔画乱排。"""
    s = str(slot or "")
    out = []
    if "顶" in s:
        out.append(-1)
    for ch in s:
        if ch in CN_NUM:
            out.append(CN_NUM[ch])
        elif ch.isdigit():
            out.append(int(ch))
    if "底" in s:
        out.append(90)
    # 左 < 中 < 右
    for ch, n in (("左", 1), ("中", 2), ("右", 3)):
        if ch in s:
            out.append(n)
    return out


def cab_tokens(cab):
    s = str(cab or "").lower()
    out = [s]
    n = cab_index(s)
    if n < 11:
        cn = "零一二三四五六七八九十"[n]
        out += ["%d号柜" % n, "%d号" % n, "柜%d" % n, str(n), cn + "号柜", cn + "号"]
    return " ".join(out)


def haystack(it):
    return " ".join([it.get(k, "") for k in FIELDS] + [cab_tokens(it.get("cabinet"))]).lower()


class Editor(tk.Toplevel):
    def __init__(self, app, item=None):
        super().__init__(app.root)
        self.app = app
        self.item = item or {}
        self.title("编辑材料" if item else "新增材料")
        self.resizable(False, False)
        self.transient(app.root)
        self.grab_set()
        self.vars = {}
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        body.columnconfigure(3, weight=1)

        specs = [
            ("name", 0, 3), ("alias", 1, 3),
            ("cas", 2, 1), ("spec", 2, 1),
            ("category", 3, 1), ("cabinet", 3, 1),
            ("slot", 4, 1), ("qty", 4, 1),
            ("note", 5, 3),
        ]
        row = 0
        for i, (key, grid_row, span) in enumerate(specs):
            col = 0 if i % 2 == 0 or span == 3 else 2
            row = grid_row
            text = LABELS[key] + (" *" if key == "name" else "")
            ttk.Label(body, text=text).grid(row=row, column=col, sticky="w", padx=(0, 8), pady=6)
            var = tk.StringVar(value=str(self.item.get(key, "")))
            self.vars[key] = var
            if key in ("category", "cabinet"):
                values = CATS if key == "category" else CABINETS
                ent = ttk.Combobox(body, textvariable=var, values=values)
            else:
                ent = ttk.Entry(body, textvariable=var)
            ent.grid(row=row, column=col + 1, columnspan=span, sticky="ew", pady=6)
        ttk.Label(body, text="提示：别名里可以写化学式、英文名、俗名，搜索时都能命中。",
                  foreground="#666").grid(row=6, column=0, columnspan=4, sticky="w", pady=(8, 0))

        btns = ttk.Frame(body)
        btns.grid(row=7, column=0, columnspan=4, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(btns, text="保存", command=self.save).pack(side="right")

        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Return>", lambda e: self.save())
        self.wait_visibility()
        self.focus_force()

    def save(self):
        data = {k: self.vars[k].get().strip() for k in FIELDS}
        if not data["name"]:
            messagebox.showwarning(APP_NAME, "名称不能为空", parent=self)
            return
        self.app.upsert(data, self.item.get("id"))
        self.destroy()


class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_NAME)
        root.geometry("1040x680")
        root.minsize(820, 520)
        self.items = load_items()
        self.q = tk.StringVar()
        self.cab = tk.StringVar(value="全部柜位")
        self.cat = tk.StringVar(value="全部类型")
        self.sort = tk.StringVar(value="按柜位排")
        self._build()
        self.refresh()

    # ---------- 界面 ----------
    def _build(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(".", font=("Microsoft YaHei UI", 10))
        style.configure("Treeview", rowheight=30, font=("Microsoft YaHei UI", 10))
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 13, "bold"))
        style.configure("Hint.TLabel", foreground="#666")

        outer = ttk.Frame(self.root, padding=(16, 12, 16, 0))
        outer.pack(fill="both", expand=True)

        head = ttk.Frame(outer)
        head.pack(fill="x")
        ttk.Label(head, text="试剂柜速查", style="Title.TLabel").pack(side="left")
        ttk.Button(head, text="+ 新增", command=lambda: self.edit(None)).pack(side="right", padx=(6, 0))
        ttk.Button(head, text="统计", command=self.show_stats).pack(side="right", padx=(6, 0))
        ttk.Button(head, text="导入 / 导出", command=self.data_menu).pack(side="right")
        ttk.Label(head, text="双击一行可以修改", style="Hint.TLabel").pack(side="left", padx=(12, 0))

        search = ttk.Frame(outer)
        search.pack(fill="x", pady=(10, 6))
        ttk.Label(search, text="搜索").pack(side="left", padx=(0, 8))
        ent = ttk.Entry(search, textvariable=self.q, font=("Microsoft YaHei UI", 11))
        ent.pack(side="left", fill="x", expand=True)
        ent.bind("<KeyRelease>", lambda e: self.refresh())
        self.search_entry = ent
        ttk.Label(search, text="名称/别名/化学式/CAS/柜号", style="Hint.TLabel").pack(side="left", padx=(8, 0))

        filt = ttk.Frame(outer)
        filt.pack(fill="x", pady=(0, 8))
        ttk.Label(filt, text="柜位").pack(side="left", padx=(0, 6))
        self.cab_box = ttk.Combobox(filt, textvariable=self.cab, state="readonly", width=10)
        self.cab_box.pack(side="left")
        self.cab_box.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        ttk.Label(filt, text="类型").pack(side="left", padx=(14, 6))
        self.cat_box = ttk.Combobox(filt, textvariable=self.cat, state="readonly", width=12)
        self.cat_box.pack(side="left")
        self.cat_box.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        ttk.Label(filt, text="排序").pack(side="left", padx=(14, 6))
        self.sort_box = ttk.Combobox(filt, textvariable=self.sort, state="readonly", width=10,
                                     values=["按柜位排", "按名称排", "按类型排", "最近更新"])
        self.sort_box.pack(side="left")
        self.sort_box.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        ttk.Button(filt, text="清空条件", command=self.clear_filters).pack(side="right")

        table = ttk.Frame(outer)
        table.pack(fill="both", expand=True)
        cols = ("name", "loc", "alias", "spec", "category", "qty")
        self.tree = ttk.Treeview(table, columns=cols, show="headings", selectmode="browse")
        heads = {"name": ("名称", 180), "loc": ("位置", 190), "alias": ("别名 / 化学式", 260),
                 "spec": ("规格", 150), "category": ("类型", 100), "qty": ("数量", 80)}
        for c in cols:
            text, width = heads[c]
            self.tree.heading(c, text=text)
            self.tree.column(c, width=width, anchor="w", stretch=(c in ("alias", "name")))
        vsb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", lambda e: self.edit_selected())
        self.tree.bind("<Delete>", lambda e: self.delete_selected())
        self.tree.bind("<Return>", lambda e: self.edit_selected())
        self.tree.tag_configure("odd", background="#fafafa")

        bottom = ttk.Frame(self.root, padding=(16, 6, 16, 8))
        bottom.pack(fill="x")
        self.status = ttk.Label(bottom, text="", style="Hint.TLabel")
        self.status.pack(side="left")
        ttk.Button(bottom, text="删除所选", command=self.delete_selected).pack(side="right")
        ttk.Button(bottom, text="修改所选", command=self.edit_selected).pack(side="right", padx=(0, 6))

        self.root.bind("<Control-n>", lambda e: self.edit(None))
        self.root.bind("<Control-f>", lambda e: (self.search_entry.focus_set(), self.search_entry.select_range(0, "end")))
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.search_entry.focus_set()

    # ---------- 数据 ----------
    def save(self):
        try:
            write_items([{k: it.get(k, "") for k in FIELDS} | {"id": it.get("id", ""), "updatedAt": it.get("updatedAt", 0)}
                         for it in self.items])
            self.status.configure(text="已保存 %s    %d 条    %s" % (now_text(), len(self.items), DATA_FILE))
        except Exception as exc:
            messagebox.showerror(APP_NAME, "保存失败：%r\n\n请检查程序所在文件夹是否可写。" % (exc,))

    def filtered(self):
        ts = [t for t in self.q.get().strip().lower().split() if t]
        cab, cat = self.cab.get(), self.cat.get()
        out = []
        for it in self.items:
            if cab != "全部柜位" and it.get("cabinet") != cab:
                continue
            if cat != "全部类型" and it.get("category") != cat:
                continue
            if ts:
                h = haystack(it)
                if not all(t in h for t in ts):
                    continue
            out.append(it)
        key = self.sort.get()
        if key == "按名称排":
            out.sort(key=lambda x: x.get("name", ""))
        elif key == "按类型排":
            out.sort(key=lambda x: (x.get("category", ""), x.get("name", "")))
        elif key == "最近更新":
            out.sort(key=lambda x: -int(x.get("updatedAt") or 0))
        else:
            out.sort(key=lambda x: (cab_index(x.get("cabinet")), slot_key(x.get("slot")), x.get("name", "")))
        return out

    def refresh(self):
        cabs, cats = {"全部柜位"}, {"全部类型"}
        for it in self.items:
            if it.get("cabinet"):
                cabs.add(it["cabinet"])
            if it.get("category"):
                cats.add(it["category"])
        self.cab_box["values"] = sorted(cabs, key=cab_index)
        self.cat_box["values"] = sorted(cats, key=lambda c: (CATS.index(c) if c in CATS else 99, c))
        if self.cab.get() not in self.cab_box["values"]:
            self.cab.set("全部柜位")
        if self.cat.get() not in self.cat_box["values"]:
            self.cat.set("全部类型")

        self.tree.delete(*self.tree.get_children())
        rows = self.filtered()
        for i, it in enumerate(rows):
            loc = " · ".join(x for x in (it.get("cabinet"), it.get("slot")) if x) or "（未填位置）"
            self.tree.insert("", "end", iid=it["id"], tags=("odd",) if i % 2 else (),
                             values=(it.get("name", ""), loc, it.get("alias", ""),
                                     it.get("spec", ""), it.get("category", ""), it.get("qty", "")))
        total = len(self.items)
        shown = len(rows)
        self.status.configure(text=("筛出 %d 条 / 共 %d 条" % (shown, total)) if shown != total else ("共 %d 条" % total))

    def selected_item(self):
        sel = self.tree.selection()
        if not sel:
            return None
        for it in self.items:
            if it["id"] == sel[0]:
                return it
        return None

    def edit(self, item):
        Editor(self, item)

    def edit_selected(self):
        it = self.selected_item()
        if not it:
            messagebox.showinfo(APP_NAME, "先在列表里选中一行")
            return
        self.edit(it)

    def upsert(self, data, item_id=None):
        if item_id:
            for it in self.items:
                if it["id"] == item_id:
                    it.update(data)
                    it["updatedAt"] = int(datetime.now().timestamp() * 1000)
                    break
        else:
            data = dict(data)
            data["id"] = new_id()
            data["updatedAt"] = int(datetime.now().timestamp() * 1000)
            self.items.append(data)
        self.save()
        self.refresh()
        for iid in self.tree.get_children():
            if iid == (item_id or data.get("id")):
                self.tree.selection_set(iid)
                self.tree.see(iid)

    def delete_selected(self):
        it = self.selected_item()
        if not it:
            return
        if not messagebox.askyesno(APP_NAME, "确定删除「%s」？" % it.get("name", "")):
            return
        self.items = [x for x in self.items if x["id"] != it["id"]]
        self.save()
        self.refresh()

    def clear_filters(self):
        self.q.set("")
        self.cab.set("全部柜位")
        self.cat.set("全部类型")
        self.refresh()

    # ---------- 统计 ----------
    def show_stats(self):
        win = tk.Toplevel(self.root)
        win.title("统计")
        win.geometry("520x560")
        win.transient(self.root)
        frame = ttk.Frame(win, padding=18)
        frame.pack(fill="both", expand=True)
        total = len(self.items)
        ttk.Label(frame, text="共 %d 条材料" % total, style="Title.TLabel").pack(anchor="w")
        for key, title in (("cabinet", "按柜子分布"), ("category", "按类型分布")):
            counts = {}
            for it in self.items:
                counts[it.get(key) or "（未填）"] = counts.get(it.get(key) or "（未填）", 0) + 1
            ttk.Label(frame, text="\n" + title, font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w")
            if not counts:
                ttk.Label(frame, text="（暂无数据）", foreground="#666").pack(anchor="w")
                continue
            top = max(counts.values()) or 1
            for name, n in sorted(counts.items(), key=lambda x: (cab_index(x[0]) if key == "cabinet" else -x[1], x[0])):
                bar = "█" * max(1, round(n / top * 24))
                row = ttk.Frame(frame)
                row.pack(fill="x", pady=1)
                ttk.Label(row, text="%-10s" % name, width=12).pack(side="left")
                ttk.Label(row, text=bar, foreground="#0d9488").pack(side="left")
                ttk.Label(row, text=" %d 条" % n, foreground="#666").pack(side="left")
        ttk.Button(frame, text="关闭", command=win.destroy).pack(anchor="e", pady=(16, 0))

    # ---------- 导入导出 ----------
    def data_menu(self):
        win = tk.Toplevel(self.root)
        win.title("导入 / 导出")
        win.geometry("440x360")
        win.transient(self.root)
        f = ttk.Frame(win, padding=18)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="数据文件：", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w")
        ttk.Label(f, text=DATA_FILE, wraplength=370, foreground="#666").pack(anchor="w", pady=(0, 12))
        for text, cmd in (
            ("导出备份（JSON）", self.export_json),
            ("从备份导入（JSON）", self.import_json),
            ("导出表格（CSV，Excel 可开）", self.export_csv),
            ("从 Excel / CSV 批量导入", self.import_table),
            ("导出导入模板（CSV）", self.export_template),
        ):
            ttk.Button(f, text=text, command=cmd).pack(fill="x", pady=3)
        ttk.Label(f, text="支持 .xlsx 和 .csv，会自动识别列名（名称/品名、柜号/柜子、位置等）。\n"
                          "「位置」写成“一号柜 二层 左”也能自动拆开。",
                  foreground="#666", wraplength=370).pack(anchor="w", pady=(10, 0))
        ttk.Button(f, text="关闭", command=win.destroy).pack(anchor="e", pady=(8, 0))

    def export_template(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="试剂柜导入模板.csv",
                                            filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["名称", "别名", "CAS号", "规格", "类型", "柜号", "位置", "数量", "备注"])
            w.writerow(["氯化钠", "NaCl 食盐", "7647-14-5", "AR 500g", "无机盐", "一号柜", "二层 左", "2 瓶", ""])
        messagebox.showinfo(APP_NAME, "模板已保存到：\n" + path + "\n\n按这个表头填好，再用「从 Excel / CSV 批量导入」。")


    def export_json(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", initialfile="试剂柜备份.json",
                                            filetypes=[("JSON", "*.json")])
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"app": "reagent-locator", "items": self.items}, f, ensure_ascii=False, indent=2)
        messagebox.showinfo(APP_NAME, "已导出到：\n" + path)

    def import_json(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            arr = raw.get("items") if isinstance(raw, dict) else raw
            if not isinstance(arr, list):
                raise ValueError("格式不对")
            self.merge(arr)
            messagebox.showinfo(APP_NAME, "导入完成，当前共 %d 条" % len(self.items))
        except Exception as exc:
            messagebox.showerror(APP_NAME, "导入失败：%r" % (exc,))

    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="试剂柜清单.csv",
                                            filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["名称", "别名", "CAS号", "规格", "类型", "柜号", "位置", "数量", "备注"])
            for it in self.filtered():
                w.writerow([it.get(k, "") for k in FIELDS])
        messagebox.showinfo(APP_NAME, "已导出到：\n" + path)

    def import_table(self):
        path = filedialog.askopenfilename(
            title="选择要导入的 Excel 或 CSV",
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"),
                       ("Excel 工作簿", "*.xlsx *.xlsm"),
                       ("CSV 文件", "*.csv")])
        if not path:
            return
        try:
            rows = read_table(path)
            found = detect_header(rows)
            if not found:
                raise ValueError("没找到「名称」列。\n\n表头里需要有「名称 / 品名 / 试剂名称」之类的列名；"
                                 "如果第一行是标题，程序会自动往下找 10 行。")
            header_row, mapping = found
            items = rows_to_items(rows, header_row, mapping)
            if not items:
                raise ValueError("没有读到有效数据行（「名称」为空的行会被跳过）。")

            heads = rows[header_row]
            desc = "\n".join("    %s  ←  %s" % (LABELS[k], heads[i]) for k, i in mapping.items())
            if not messagebox.askyesno(
                    APP_NAME,
                    "识别到 %d 行数据。\n\n列对应关系：\n%s\n\n"
                    "导入时按「名称 + 柜号 + 位置」去重，重复的会更新。继续吗？" % (len(items), desc)):
                return

            added, updated = self.merge(items)
            messagebox.showinfo(APP_NAME, "导入完成：新增 %d 条，更新 %d 条，当前共 %d 条。"
                                % (added, updated, len(self.items)))
        except Exception as exc:
            messagebox.showerror(APP_NAME, "导入失败：\n%s" % exc)

    def merge(self, arr):
        index = {(x.get("name"), x.get("cabinet"), x.get("slot")): x for x in self.items}
        added = updated = 0
        for raw in arr:
            d = {k: str(raw.get(k, "") or "").strip() for k in FIELDS}
            if not d["name"]:
                continue
            key = (d["name"], d["cabinet"], d["slot"])
            if key in index:
                index[key].update(d)
                updated += 1
            else:
                d["id"] = new_id()
                d["updatedAt"] = int(datetime.now().timestamp() * 1000)
                self.items.append(d)
                index[key] = d
                added += 1
        self.save()
        self.refresh()
        return added, updated

    def on_close(self):
        self.save()
        self.root.destroy()


def seed_items():
    rows = [
        ("氯化钠", "NaCl sodium chloride 食盐", "7647-14-5", "AR 500g", "无机盐", "一号柜", "二层 左", "2 瓶", ""),
        ("硫酸铜", "CuSO4 五水硫酸铜 胆矾", "7758-98-7", "AR 500g", "无机盐", "一号柜", "三层 中", "1 瓶", ""),
        ("无水乙醇", "乙醇 ethanol C2H5OH", "64-17-5", "AR 500mL", "有机溶剂", "二号柜", "一层 右", "3 瓶", "易燃"),
        ("丙酮", "acetone C3H6O", "67-64-1", "AR 500mL", "有机溶剂", "二号柜", "一层 左", "2 瓶", "易燃，通风"),
        ("盐酸", "HCl 氢氯酸", "7647-01-0", "36%~38% 500mL", "酸碱", "三号柜", "底层 通风", "1 瓶", "腐蚀性"),
        ("氢氧化钠", "NaOH 烧碱 火碱", "1310-73-2", "AR 500g", "酸碱", "三号柜", "二层 右", "2 瓶", "腐蚀性"),
        ("过氧化氢", "H2O2 双氧水", "7722-84-1", "30% 500mL", "氧化剂", "三号柜", "底层 避光", "1 瓶", "避光"),
        ("酚酞", "phenolphthalein", "77-09-8", "1% 乙醇溶液 100mL", "指示剂", "四号柜", "一层 左", "1 瓶", ""),
        ("琼脂糖", "agarose", "9012-36-6", "生物级 100g", "生物试剂", "四号柜", "二层 右", "1 瓶", ""),
    ]
    out = []
    for i, r in enumerate(rows):
        d = dict(zip(FIELDS, r))
        d["id"] = "seed%d" % i
        d["updatedAt"] = int(datetime.now().timestamp() * 1000) - i * 1000
        out.append(d)
    return out


_MUTEX = None


def focus_existing_window():
    """把已经在运行的窗口切到前台。"""
    import ctypes
    u = ctypes.windll.user32
    hwnd = u.FindWindowW(None, APP_NAME)
    if hwnd:
        u.ShowWindow(hwnd, 9)
        u.SetForegroundWindow(hwnd)
        return True
    return False


def acquire_single_instance():
    """命名互斥体，保证同时只有一个实例在写 data.json。"""
    import ctypes
    kernel32 = ctypes.windll.kernel32
    kernel32.CreateMutexW.restype = ctypes.c_void_p
    global _MUTEX
    _MUTEX = kernel32.CreateMutexW(None, False, "Local\\ReagentCabinetSingleInstance")
    return kernel32.GetLastError() != 183  # ERROR_ALREADY_EXISTS


def main():
    if not acquire_single_instance():
        import time
        for _ in range(10):
            if focus_existing_window():
                break
            time.sleep(0.3)
        return
    daily_backup()
    root = tk.Tk()
    app = App(root)
    # 首次运行：自动写入示例，省得空窗口无从下手
    if not app.items:
        app.items = seed_items()
        app.save()
        app.refresh()
    root.mainloop()


def cli_import_test(src, out):
    """开发自检用：解析一个 Excel/CSV，把识别结果写到 out（窗口程序没有 stdout）。"""
    rows = read_table(src)
    found = detect_header(rows)
    if not found:
        raise SystemExit("no header")
    items = rows_to_items(rows, found[0], found[1])
    with open(out, "w", encoding="utf-8") as f:
        json.dump({"count": len(items), "first": items[0] if items else None,
                   "columns": {k: rows[found[0]][i] for k, i in found[1].items()}},
                  f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    if len(sys.argv) >= 4 and sys.argv[1] == "--import-test":
        cli_import_test(sys.argv[2], sys.argv[3])
    else:
        main()

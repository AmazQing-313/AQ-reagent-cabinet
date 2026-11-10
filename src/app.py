# -*- coding: utf-8 -*-
"""
试剂柜速查 · 桌面版
- 纯本地窗口程序，不需要浏览器、不联网。
- 数据保存在 exe 同目录的 data.json。
"""
import csv
import json
import os
import shutil
import sys
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
            row["id"] = str(it.get("id") or ("i" + str(len(out))))
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
            data["id"] = "i%d" % (int(datetime.now().timestamp() * 1000) % 10 ** 12)
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
        win.geometry("420x300")
        win.transient(self.root)
        f = ttk.Frame(win, padding=18)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="数据文件：", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w")
        ttk.Label(f, text=DATA_FILE, wraplength=360, foreground="#666").pack(anchor="w", pady=(0, 12))
        for text, cmd in (
            ("导出备份（JSON）", self.export_json),
            ("从备份导入（JSON）", self.import_json),
            ("导出表格（CSV，Excel 可开）", self.export_csv),
            ("从 CSV 批量导入", self.import_csv),
        ):
            ttk.Button(f, text=text, command=cmd).pack(fill="x", pady=3)
        ttk.Label(f, text="CSV 表头：名称,别名,CAS号,规格,类型,柜号,位置,数量,备注",
                  foreground="#666", wraplength=360).pack(anchor="w", pady=(10, 0))
        ttk.Button(f, text="关闭", command=win.destroy).pack(anchor="e", pady=(8, 0))

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

    def import_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                rows = list(csv.reader(f))
            if len(rows) < 2:
                raise ValueError("CSV 里没有数据行")
            head = [h.strip() for h in rows[0]]
            mapping = {}
            for idx, name in enumerate(head):
                for key, label in LABELS.items():
                    if name in (label, label.replace(" / ", "")):
                        mapping[key] = idx
            if "name" not in mapping:
                raise ValueError("没有找到「名称」列")
            arr = []
            for r in rows[1:]:
                if not any(c.strip() for c in r):
                    continue
                d = {k: (r[i].strip() if i < len(r) else "") for k, i in mapping.items()}
                if d.get("name"):
                    arr.append(d)
            self.merge(arr)
            messagebox.showinfo(APP_NAME, "导入完成，当前共 %d 条" % len(self.items))
        except Exception as exc:
            messagebox.showerror(APP_NAME, "导入失败：%r" % (exc,))

    def merge(self, arr):
        index = {(x.get("name"), x.get("cabinet"), x.get("slot")): x for x in self.items}
        for raw in arr:
            d = {k: str(raw.get(k, "") or "").strip() for k in FIELDS}
            if not d["name"]:
                continue
            key = (d["name"], d["cabinet"], d["slot"])
            if key in index:
                index[key].update(d)
            else:
                d["id"] = "i%d" % (int(datetime.now().timestamp() * 1000) % 10 ** 12)
                d["updatedAt"] = int(datetime.now().timestamp() * 1000)
                self.items.append(d)
                index[key] = d
        self.save()
        self.refresh()

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


if __name__ == "__main__":
    main()

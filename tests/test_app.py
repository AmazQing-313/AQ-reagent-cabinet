# -*- coding: utf-8 -*-
"""桌面版逻辑自检：搜索、筛选、新增、去重、落盘。"""
import json
import os
import shutil
import sys
import tempfile
import tkinter as tk

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))
import app as appmod  # noqa: E402

tmp = os.path.join(ROOT, "tests", "_tmp")
shutil.rmtree(tmp, ignore_errors=True)
os.makedirs(tmp)
appmod.DATA_FILE = os.path.join(tmp, "data.json")
appmod.BACKUP_DIR = os.path.join(tmp, "备份")
shutil.copy2(os.path.join(ROOT, "examples", "data.sample.json"), appmod.DATA_FILE)

passed = failed = 0


def ok(name, cond, extra=""):
    global passed, failed
    if cond:
        passed += 1
        print("  PASS  " + name)
    else:
        failed += 1
        print("  FAIL  " + name + ("  → " + str(extra) if extra else ""))


root = tk.Tk()
root.withdraw()
app = appmod.App(root)

print("1. 加载")
ok("载入 9 条", len(app.items) == 9, len(app.items))
ok("表格渲染 9 行", len(app.tree.get_children()) == 9)

print("2. 搜索")
for kw, expect in (("氯化钠", 1), ("NaCl", 1), ("1号柜", 2), ("7647", 2), ("乙醇 HCl", 0)):
    app.q.set(kw)
    app.refresh()
    n = len(app.tree.get_children())
    ok("搜索「%s」→ %d 条" % (kw, expect), n == expect, n)

app.q.set("氯化钠")
app.refresh()
vals = app.tree.item(app.tree.get_children()[0], "values")
ok("位置列正确", vals[1] == "一号柜 · 二层 左", vals)
ok("别名可显示", "NaCl" in vals[2], vals)

app.q.set("")
app.refresh()

print("3. 筛选")
app.cab.set("二号柜")
app.refresh()
ok("柜位筛选 2 条", len(app.tree.get_children()) == 2, len(app.tree.get_children()))
app.cab.set("全部柜位")
app.cat.set("酸碱")
app.refresh()
ok("类型筛选 2 条", len(app.tree.get_children()) == 2)
app.cat.set("全部类型")
app.refresh()

print("4. 新增与落盘")
app.upsert({"name": "正庚烷", "alias": "heptane", "cas": "142-82-5", "spec": "AR 500mL",
            "category": "有机溶剂", "cabinet": "二号柜", "slot": "三层", "qty": "1 瓶", "note": ""})
ok("新增后 10 条", len(app.items) == 10, len(app.items))
raw = json.load(open(appmod.DATA_FILE, encoding="utf-8"))
ok("已写入 data.json", raw["items"][-1]["name"] == "正庚烷", raw["items"][-1]["name"])
ok("文件是 UTF-8 中文", "正庚烷" in open(appmod.DATA_FILE, encoding="utf-8").read())

print("5. 导入去重")
before = len(app.items)
app.merge([{"name": "正庚烷", "cabinet": "二号柜", "slot": "三层", "qty": "2 瓶"}])
ok("同一条更新不重复", len(app.items) == before, "%d vs %d" % (len(app.items), before))
ok("更新生效", [x for x in app.items if x["name"] == "正庚烷"][0]["qty"] == "2 瓶")

print("6. 统计窗口")
try:
    app.show_stats()
    root.update()
    ok("统计窗口能打开", True)
except Exception as exc:
    ok("统计窗口能打开", False, repr(exc))

print("7. 关闭保存")
app.on_close()
ok("窗口关闭后文件仍在", os.path.exists(appmod.DATA_FILE))

print("\n" + "=" * 44)
print("通过 %d 项，失败 %d 项" % (passed, failed))
sys.exit(1 if failed else 0)

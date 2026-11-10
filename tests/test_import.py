# -*- coding: utf-8 -*-
"""Excel / CSV 导入逻辑自检。"""
import os
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))
import app as appmod  # noqa: E402

passed = failed = 0


def ok(name, cond, extra=""):
    global passed, failed
    if cond:
        passed += 1
        print("  PASS  " + name)
    else:
        failed += 1
        print("  FAIL  " + name + ("  → " + str(extra) if extra else ""))


print("1. 表头归一化")
ok("「名称（必填）」→ 名称", appmod.normalize_header("名称（必填）") == "名称")
ok("「CAS No.」→ casno", appmod.normalize_header("CAS No.") == "casno")
ok("全角空格被去掉", appmod.normalize_header("　柜号　") == "柜号")

print("2. 列名识别（各种写法）")
m = appmod.map_columns(["试剂名称", "存放柜", "具体位置", "包装", "数量", "英文名"])
ok("名称←试剂名称", m.get("name") == 0, m)
ok("柜号←存放柜", m.get("cabinet") == 1, m)
ok("位置←具体位置", m.get("slot") == 2, m)
ok("规格←包装", m.get("spec") == 3, m)
ok("数量←数量", m.get("qty") == 4, m)
ok("别名←英文名", m.get("alias") == 5, m)
ok("未知列被忽略", "note" not in m, m)

print("3. 表头行定位（前面有标题行）")
rows = [
    ["实验室试剂清单", "", ""],
    ["制表人：张三", "", ""],
    ["名称", "柜号", "位置"],
    ["氯化钠", "一号柜", "二层 左"],
]
found = appmod.detect_header(rows)
ok("跳过标题行找到第 3 行", found is not None and found[0] == 2, found and found[0])
ok("映射出名称列", found and found[1].get("name") == 0)

print("4. 位置拆分")
ok("「一号柜 二层 左」拆分", appmod.split_location("", "一号柜 二层 左") == ("一号柜", "二层 左"),
   appmod.split_location("", "一号柜 二层 左"))
ok("「2号柜三层」拆分", appmod.split_location("", "2号柜三层") == ("2号柜", "三层"),
   appmod.split_location("", "2号柜三层"))
ok("已分列时不拆", appmod.split_location("二号柜", "一层") == ("二号柜", "一层"))

print("5. 行转条目")
items = appmod.rows_to_items(rows, 2, {"name": 0, "cabinet": 1, "slot": 2})
ok("得到 1 条", len(items) == 1, len(items))
ok("名称正确", items[0]["name"] == "氯化钠")
ok("位置被拆分", items[0]["cabinet"] == "一号柜" and items[0]["slot"] == "二层 左",
   (items[0]["cabinet"], items[0]["slot"]))

print("6. 读 CSV")
tmp = os.path.join(ROOT, "tests", "_tmp_import")
shutil.rmtree(tmp, ignore_errors=True)
os.makedirs(tmp)
csv_path = os.path.join(tmp, "list.csv")
with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
    f.write("名称,柜号,位置,规格\n氯化钠,一号柜,二层 左,AR 500g\n盐酸,三号柜,底层,500mL\n")
rows_csv = appmod.read_table(csv_path)
found = appmod.detect_header(rows_csv)
ok("CSV 读到表头", found is not None and found[0] == 0)
ok("CSV 转出 2 条", len(appmod.rows_to_items(rows_csv, found[0], found[1])) == 2)

print("7. 读 Excel（标准库解析，含标题行、合并写法、空行）")
xlsx_path = os.path.join(ROOT, "examples", "import-example.xlsx")
rows_x = appmod.read_table(xlsx_path)
found = appmod.detect_header(rows_x)
ok("Excel 找到表头（跳过标题与空行）", found is not None and found[0] == 2, found and found[0])
items_x = appmod.rows_to_items(rows_x, found[0], found[1])
ok("Excel 转出 4 条（跳过空行）", len(items_x) == 4, len(items_x))
ok("合并位置被拆开", items_x[0]["cabinet"] == "一号柜" and items_x[0]["slot"] == "二层 左",
   (items_x[0]["cabinet"], items_x[0]["slot"]))
ok("数字单元格不成 2.0", items_x[0]["qty"] == "2", items_x[0]["qty"])
ok("第二行正常", items_x[1]["name"] == "无水乙醇" and items_x[1]["cabinet"] == "二号柜")
ok("CAS 列被识别", items_x[0]["cas"] == "7647-14-5", items_x[0]["cas"])

print("8. 合并去重（走 App.merge）")
import tkinter as tk
appmod.DATA_FILE = os.path.join(tmp, "data.json")
appmod.BACKUP_DIR = os.path.join(tmp, "备份")
with open(appmod.DATA_FILE, "w", encoding="utf-8") as f:
    f.write('{"app":"reagent-locator","items":[]}')
root = tk.Tk()
root.withdraw()
a = appmod.App(root)
added, updated = a.merge(items_x)
ok("首次新增 4 条", (added, updated) == (4, 0), (added, updated))
added, updated = a.merge(items_x)
ok("再次导入变成更新", (added, updated) == (0, 4), (added, updated))
ok("没有重复条目", len(a.items) == 4, len(a.items))
root.destroy()

shutil.rmtree(tmp, ignore_errors=True)

print("\n" + "=" * 44)
print("通过 %d 项，失败 %d 项" % (passed, failed))
sys.exit(1 if failed else 0)

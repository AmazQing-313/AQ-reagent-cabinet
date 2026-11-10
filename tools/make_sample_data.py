# -*- coding: utf-8 -*-
"""生成 examples/data.sample.json（9 条常见试剂）。"""
import json
import os
import time

ROWS = [
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

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
now = int(time.time() * 1000)
items = []
for i, r in enumerate(ROWS):
    items.append({
        "id": "seed%d" % i, "name": r[0], "alias": r[1], "cas": r[2], "spec": r[3],
        "category": r[4], "cabinet": r[5], "slot": r[6], "qty": r[7], "note": r[8],
        "updatedAt": now - i * 1000,
    })
out = os.path.join(ROOT, "examples", "data.sample.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump({"app": "reagent-locator", "savedAt": None, "items": items}, f,
              ensure_ascii=False, indent=2)
print("written", out, len(items), "items")

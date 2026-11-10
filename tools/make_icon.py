# -*- coding: utf-8 -*-
"""生成 exe 图标：圆角深青底 + 白色锥形瓶。"""
import os
from PIL import Image, ImageDraw

S = 1024
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

# 背景：圆角方块，轻微渐变
d.rounded_rectangle([0, 0, S - 1, S - 1], radius=int(S * 0.22), fill=(13, 148, 136, 255))

# 顶部高光
d.rounded_rectangle([int(S * 0.06), int(S * 0.05), int(S * 0.94), int(S * 0.42)],
                    radius=int(S * 0.18), fill=(20, 170, 155, 90))

# 锥形瓶：瓶颈 + 瓶身
neck_w = int(S * 0.20)
cx = S // 2
top = int(S * 0.20)
neck_bottom = int(S * 0.42)
body_bottom = int(S * 0.84)
left = int(S * 0.22)
right = int(S * 0.78)

flask = [
    (cx - neck_w // 2, top),
    (cx + neck_w // 2, top),
    (cx + neck_w // 2, neck_bottom),
    (right, body_bottom - int(S * 0.05)),
    (right, body_bottom),
    (left, body_bottom),
    (left, body_bottom - int(S * 0.05)),
    (cx - neck_w // 2, neck_bottom),
]
d.polygon(flask, fill=(255, 255, 255, 255))

# 瓶口
d.rounded_rectangle([cx - neck_w // 2 - int(S * 0.03), top - int(S * 0.035),
                     cx + neck_w // 2 + int(S * 0.03), top + int(S * 0.03)],
                    radius=int(S * 0.02), fill=(255, 255, 255, 255))

# 液体：瓶身下部
liquid_top = int(S * 0.63)
d.polygon([
    (int(S * 0.31), liquid_top),
    (int(S * 0.69), liquid_top),
    (int(S * 0.76), body_bottom - int(S * 0.02)),
    (int(S * 0.24), body_bottom - int(S * 0.02)),
], fill=(13, 148, 136, 255))
d.rectangle([int(S * 0.24), liquid_top - int(S * 0.012), int(S * 0.76), liquid_top + int(S * 0.012)],
            fill=(45, 212, 191, 255))

# 高光条
d.rounded_rectangle([int(S * 0.33), int(S * 0.30), int(S * 0.37), int(S * 0.58)],
                    radius=int(S * 0.02), fill=(255, 255, 255, 120))

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
img.save(out, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print("icon ->", out, os.path.getsize(out), "bytes")

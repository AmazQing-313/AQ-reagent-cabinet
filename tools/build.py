# -*- coding: utf-8 -*-
"""用 PyInstaller 打包 Windows 版（需先 pip install pyinstaller）。"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--windowed", "--onedir",
        "--name", "试剂柜速查",
        "--icon", os.path.join(ROOT, "assets", "icon.ico"),
        os.path.join(ROOT, "src", "app.py"),
    ]
    print(" ".join(cmd))
    return subprocess.call(cmd, cwd=ROOT)


if __name__ == "__main__":
    sys.exit(main())

# -*- coding: utf-8 -*-
"""用 PyInstaller 打包 Windows 版（需先 pip install pyinstaller）。"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def ensure_utf8_stdio():
    """Windows 控制台可能是 cp1252 / cp936，直接打印中文会抛 UnicodeEncodeError。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def main():
    ensure_utf8_stdio()
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--windowed", "--onedir",
        "--name", "试剂柜速查",
        "--icon", os.path.join(ROOT, "assets", "icon.ico"),
        os.path.join(ROOT, "src", "app.py"),
    ]
    print("building:", " ".join(cmd))
    code = subprocess.call(cmd, cwd=ROOT)
    print("PyInstaller exit code:", code)
    return code


if __name__ == "__main__":
    sys.exit(main())

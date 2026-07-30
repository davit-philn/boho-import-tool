"""
main.py — BOHO BOM2BRAVO entry point.
Thay thế main_v6.py (monolith). Khởi chạy: python main.py
"""
import sys
import os
import json

# Đảm bảo Tools/ nằm trong sys.path khi chạy trực tiếp
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import customtkinter as ctk
from services.utils import SETTINGS_FILE

# Đọc theme đã lưu; nếu chưa có → mặc định Dark
_saved_mode = "Dark"
try:
    with open(SETTINGS_FILE, encoding="utf-8") as _sf:
        _cfg = json.load(_sf)
        _saved_mode = _cfg.get("theme", "Dark")
except Exception:
    pass

ctk.set_appearance_mode(_saved_mode)
ctk.set_default_color_theme("blue")

from views.main_window import BOMToolApp

if __name__ == "__main__":
    app = BOMToolApp()
    app.mainloop()

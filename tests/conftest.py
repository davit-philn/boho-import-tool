"""
conftest.py — Thêm Tools/ vào sys.path để test có thể import services.*
"""
import sys
import os

# Cho phép: from services.bom_parser import ...
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "Tools"))

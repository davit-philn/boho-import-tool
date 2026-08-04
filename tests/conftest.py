"""
conftest.py — Pytest fixtures dùng chung cho toàn bộ test suite BOHO Import Tool.

Session-scoped fixtures:
  - mapping          : load mapping file 1 lần
  - db_conn          : kết nối SQL Server (skip nếu VPN chưa bật)
  - good_excel_path  : path đến file BOM hợp lệ trong tests/fixtures/
  - bad_excel_path   : path đến file BOM chứa dữ liệu lỗi

Function-scoped fixtures:
  - db_tx            : TRANSACTION per-test → luôn ROLLBACK sau khi test xong
"""
import sys
import os
import json
import pytest

# Cho phép: from services.xxx import ...
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "Tools"))

# ── Paths ──────────────────────────────────────────────────────────────────────
FIXTURES_DIR   = os.path.join(os.path.dirname(__file__), "fixtures")
DB_CONFIG_FILE = os.path.join(
    os.path.dirname(__file__), "..", "Tools", "config", "db_config.json"
)


# ── Pytest markers ─────────────────────────────────────────────────────────────

def pytest_configure(config):
    config.addinivalue_line(
        "markers",
        "integration: tests cần kết nối DB thật (VPN required)",
    )


# ── DB config ─────────────────────────────────────────────────────────────────

@pytest.fixture(scope="session")
def db_config():
    """Đọc db_config.json → dict. Fail rõ ràng nếu file không có."""
    if not os.path.exists(DB_CONFIG_FILE):
        pytest.skip(f"db_config.json không tìm thấy: {DB_CONFIG_FILE}")
    with open(DB_CONFIG_FILE, encoding="utf-8") as f:
        return json.load(f)


# ── DB connection (session-scoped, bị skip nếu không VPN) ────────────────────

@pytest.fixture(scope="session")
def db_conn(db_config):
    """
    Mở 1 connection SQL Server cho toàn session test.
    Tự động SKIP toàn bộ @integration tests nếu VPN không bật / DB unreachable.
    """
    import pyodbc
    cfg = db_config
    conn_str = (
        f"DRIVER={{{cfg['driver']}}};"
        f"SERVER={cfg['server']};"
        f"DATABASE={cfg['database']};"
        f"UID={cfg['username']};PWD={cfg['password']};"
        "TrustServerCertificate=yes;"
    )
    try:
        conn = pyodbc.connect(conn_str, timeout=int(cfg.get("timeout", 10)))
        conn.autocommit = True          # default; tx fixture sẽ override per-test
        yield conn
        conn.close()
    except Exception as exc:
        pytest.skip(f"DB không reach được (VPN chưa bật?): {exc}")


# ── Transaction isolation: mỗi test chạy trong 1 TX riêng rồi ROLLBACK ───────

@pytest.fixture
def db_tx(db_conn):
    """
    Fixture cho mỗi @integration test:
      - setup   : bắt đầu transaction (autocommit=False)
      - teardown: ROLLBACK dù test pass hay fail → DB không bị dirty
    """
    db_conn.autocommit = False
    yield db_conn
    try:
        db_conn.rollback()
    finally:
        db_conn.autocommit = True


# ── Mapping ───────────────────────────────────────────────────────────────────

@pytest.fixture(scope="session")
def mapping():
    """Load CK_Mapping_v5.xlsx 1 lần cho toàn session (tốn ~0.5s)."""
    from services.mapping_loader import load_mapping
    return load_mapping()


# ── Excel fixture paths ───────────────────────────────────────────────────────
#
# Đặt file Excel của bạn vào thư mục tests/fixtures/ rồi điền tên file bên dưới.
# File KHÔNG được commit lên git (chứa dữ liệu thực tế của công ty).

@pytest.fixture(scope="session")
def good_excel_path():
    """
    ── ĐIỀN TÊN FILE BOM CHUẨN CỦA BẠN VÀO ĐÂY ──────────────────────────────
    File phải pass validate layer 1 không có error (warnings OK).
    """
    fname = "bom_valid.xlsx"   # << THAY TÊN FILE THỰC TẾ CỦA BẠN
    path  = os.path.join(FIXTURES_DIR, fname)
    if not os.path.exists(path):
        pytest.skip(
            f"Fixture '{fname}' chưa có.\n"
            f"Đặt file vào: {FIXTURES_DIR}\\"
        )
    return path


@pytest.fixture(scope="session")
def bad_excel_path():
    """
    ── ĐIỀN TÊN FILE BOM LỖI CỦA BẠN VÀO ĐÂY ────────────────────────────────
    File phải trigger ít nhất 1 validation error khi validate_layer1 chạy.
    """
    fname = "bom_invalid.xlsx"   # << THAY TÊN FILE THỰC TẾ CỦA BẠN
    path  = os.path.join(FIXTURES_DIR, fname)
    if not os.path.exists(path):
        pytest.skip(
            f"Fixture '{fname}' chưa có.\n"
            f"Đặt file vào: {FIXTURES_DIR}\\"
        )
    return path

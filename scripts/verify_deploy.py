"""
scripts/verify_deploy.py — Pre-flight deployment verification.

Kiểm tra môi trường trước khi bàn giao/triển khai:
  1. Config files đủ và hợp lệ
  2. DB kết nối được
  3. Bảng B20BOM / B20BOMDetail tồn tại
  4. Mapping file đọc được và có sections cần thiết

Chạy:
    cd Tools
    python ../scripts/verify_deploy.py

Exit code 0 = OK, 1 = có lỗi cần xử lý trước khi bàn giao.
"""
import sys
import os
import json
import traceback

# Khi chạy từ Tools/ hoặc từ root
_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
sys.path.insert(0, os.path.join(_ROOT, "Tools"))

CONFIG_DIR   = os.path.join(_ROOT, "Tools", "config")
DB_CONFIG    = os.path.join(CONFIG_DIR, "db_config.json")
MAPPING_FILE = os.path.join(CONFIG_DIR, "CK_Mapping_v5.xlsx")

REQUIRED_TABLES = ["B20BOM", "B20BOMDetail"]
REQUIRED_MAPPING_SECTIONS = ["HEADER", "THDM_HEADER"]

OK   = "[  OK  ]"
FAIL = "[ FAIL ]"
WARN = "[ WARN ]"


def _sep(title=""):
    line = "─" * 60
    print(f"\n{line}")
    if title:
        print(f"  {title}")
    print(line)


def check_config_files() -> bool:
    _sep("1. Config files")
    ok = True

    for path, label in [(DB_CONFIG, "db_config.json"), (MAPPING_FILE, "CK_Mapping_v5.xlsx")]:
        if os.path.exists(path):
            size = os.path.getsize(path)
            print(f"  {OK}  {label}  ({size:,} bytes)")
        else:
            print(f"  {FAIL}  {label} — THIẾU tại {path}")
            ok = False

    # Validate db_config.json có đủ key
    if os.path.exists(DB_CONFIG):
        try:
            cfg = json.loads(open(DB_CONFIG, encoding="utf-8").read())
            required_keys = {"server", "database", "username", "password"}
            missing = required_keys - cfg.keys()
            if missing:
                print(f"  {FAIL}  db_config.json thiếu key: {missing}")
                ok = False
            else:
                print(f"  {OK}  db_config.json keys đầy đủ")
                # Warn nếu password là default
                if cfg.get("password") in ("", "Boho@2026", "password"):
                    print(f"  {WARN}  Password có vẻ là default — xem xét đổi trước bàn giao")
        except Exception as e:
            print(f"  {FAIL}  db_config.json không parse được: {e}")
            ok = False

    return ok


def check_db_connection() -> bool:
    _sep("2. DB Connection")
    try:
        import pyodbc
        cfg = json.loads(open(DB_CONFIG, encoding="utf-8").read())
        conn_str = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={cfg['server']};"
            f"DATABASE={cfg['database']};"
            f"UID={cfg['username']};"
            f"PWD={cfg['password']};"
            "Connection Timeout=8;"
        )
        conn = pyodbc.connect(conn_str, timeout=8)
        cur = conn.cursor()
        cur.execute("SELECT @@VERSION")
        ver = cur.fetchone()[0].split("\n")[0].strip()
        conn.close()
        print(f"  {OK}  Kết nối thành công: {cfg['server']}/{cfg['database']}")
        print(f"         {ver}")
        return True
    except Exception as e:
        print(f"  {FAIL}  Kết nối thất bại: {e}")
        print(f"         Kiểm tra VPN + thông tin trong db_config.json")
        return False


def check_db_tables() -> bool:
    _sep("3. Required Tables")
    try:
        import pyodbc
        cfg = json.loads(open(DB_CONFIG, encoding="utf-8").read())
        conn_str = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={cfg['server']};"
            f"DATABASE={cfg['database']};"
            f"UID={cfg['username']};"
            f"PWD={cfg['password']};"
            "Connection Timeout=8;"
        )
        conn = pyodbc.connect(conn_str, timeout=8)
        cur = conn.cursor()

        ok = True
        for tbl in REQUIRED_TABLES:
            cur.execute(
                "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ?",
                (tbl,)
            )
            count = cur.fetchone()[0]
            if count > 0:
                cur.execute(f"SELECT COUNT(*) FROM [{tbl}]")
                rows = cur.fetchone()[0]
                print(f"  {OK}  {tbl}  ({rows:,} rows hiện tại)")
            else:
                print(f"  {FAIL}  Bảng '{tbl}' KHÔNG TỒN TẠI trong DB!")
                ok = False
        conn.close()
        return ok
    except Exception as e:
        print(f"  {FAIL}  Lỗi khi kiểm tra bảng: {e}")
        return False


def check_mapping_file() -> bool:
    _sep("4. Mapping File")
    try:
        import sys as _sys
        _sys.path.insert(0, os.path.join(_ROOT, "Tools"))
        from services.mapping_loader import load_mapping
        m = load_mapping()
        ok = True
        for section in REQUIRED_MAPPING_SECTIONS:
            rows = m.get(section, [])
            if rows:
                print(f"  {OK}  Section '{section}'  ({len(rows)} rows)")
            else:
                print(f"  {FAIL}  Section '{section}' RỖNG hoặc KHÔNG TÌM THẤY")
                ok = False
        return ok
    except Exception as e:
        print(f"  {FAIL}  Không load được mapping: {e}")
        traceback.print_exc()
        return False


def check_built_exe() -> bool:
    _sep("5. Built EXE (tùy chọn)")
    exe = os.path.join(_ROOT, "Tools", "dist", "BOHO_IMPORT_BOM_THDM", "BOHO_IMPORT_BOM_THDM.exe")
    if os.path.exists(exe):
        size_mb = os.path.getsize(exe) / 1024 / 1024
        print(f"  {OK}  EXE tồn tại ({size_mb:.1f} MB): {exe}")
        return True
    else:
        print(f"  {WARN}  EXE chưa build tại {exe}")
        print(f"         Chạy: cd Tools && python build.py")
        return True  # not a hard failure


def main():
    print(f"\n{'='*60}")
    print("  BOHO IMPORT BOM/THDM — Pre-Deploy Verification")
    print(f"{'='*60}")

    if not os.path.exists(DB_CONFIG):
        print(f"\n{WARN}  Chạy script từ thư mục gốc project hoặc Tools/")

    results = {
        "Config files": check_config_files(),
    }

    # Chỉ check DB nếu config OK
    if results["Config files"]:
        db_ok = check_db_connection()
        results["DB connection"] = db_ok
        if db_ok:
            results["DB tables"] = check_db_tables()
    else:
        print(f"\n  {WARN}  Bỏ qua DB check do config bị lỗi")

    results["Mapping file"] = check_mapping_file()
    results["Built EXE"]    = check_built_exe()

    _sep("Tổng kết")
    all_ok = True
    for name, ok in results.items():
        icon = OK if ok else FAIL
        print(f"  {icon}  {name}")
        if not ok:
            all_ok = False

    if all_ok:
        print(f"\n  Tất cả kiểm tra PASS — sẵn sàng bàn giao!")
    else:
        print(f"\n  Có lỗi cần xử lý trước khi bàn giao.")

    sys.exit(0 if all_ok else 1)


if __name__ == "__main__":
    main()

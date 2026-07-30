"""
Test insert 1 dòng vào B20BOM với defaults đã chốt.
Chạy: python Tools\test_insert_bom.py

Mục đích: xác nhận cấu trúc cột + defaults đúng trước khi nối vào Excel.
"""
import json, os, sys, datetime, uuid
from pathlib import Path

BASE   = Path(__file__).parent.parent
CONFIG = BASE / "Tools" / "config" / "db_config.json"

# ── Defaults đã chốt ───────────────────────────────────────────────────────
DEFAULTS = {
    "ParentId"          : -1,
    "IsGroup"           : 0,
    "BranchCode"        : "A01",
    "DocCode"           : "BOM",
    "DocStatus"         : 4,
    "DocProcessId"      : "",
    "RowId_DocProcess"  : "",
    "ProductionProcessId": None,
    "Version"           : "1",
    "DrawingNumber"     : "",
    "Description"       : "",
    "BOMType"           : "NS",
    "ClassCode1"        : "",
    "ClassCode2"        : "",
    "ClassCode3"        : "",
    "IsActive"          : 1,
    "CreatedBy"         : 720,
    "ModifiedBy"        : 720,
    "EmployeeId"        : 1,
    "WorkProcessId"     : None,
    "ApprovalStatus"    : None,
    "ApprovedUserIdList": None,
    "UserIdList"        : None,
    "IsDraftData"       : 0,
    "ESignLayoutName"   : "",
    "PeriodId"          : None,
    "Version0"          : None,
    "MaterialSP"        : None,
    "EdgeCountAdj"      : 0,
    "TypeB"             : "",
}

def gen_row_id() -> str:
    """Sinh RowId tạm 16 ký tự: YYYYMMDD + 8 hex."""
    date_part = datetime.date.today().strftime("%Y%m%d")
    hex_part  = uuid.uuid4().hex[:8].upper()
    return date_part + hex_part        # 8 + 8 = 16 chars

def make_conn_str(cfg: dict) -> str:
    base = (
        f"DRIVER={{{cfg['driver']}}};"
        f"SERVER={cfg['server']};"
        f"DATABASE={cfg['database']};"
        "TrustServerCertificate=yes;"
    )
    if cfg.get("trusted_connection", "").lower() == "yes":
        return base + "Trusted_Connection=yes;"
    return base + f"UID={cfg['username']};PWD={cfg['password']};"

def main():
    with open(CONFIG, encoding="utf-8") as f:
        cfg = json.load(f)

    try:
        import pyodbc
    except ImportError:
        print("THIẾU: pip install pyodbc"); sys.exit(1)

    print(f"Kết nối {cfg['server']} / {cfg['database']}...", end=" ")
    conn = pyodbc.connect(make_conn_str(cfg), timeout=10)
    print("OK\n")

    now = datetime.datetime.now()

    # ── Dữ liệu test (giả lập 1 dòng BOM từ Excel) ─────────────────────────
    # Các trường này sẽ đến từ Excel khi nối thực tế
    from_excel = {
        "Code"          : "TEST-BOM-001",
        "DocDate"       : datetime.date(2025, 9, 5),
        "ItemId0"       : 19475,          # Id thực trong B20Item vừa import
        "ProductId"     : None,
        "ProductId1"    : None,
        "ParentBizDocId": None,
        "DetailRowId_SO": None,
        "EffectiveDate" : datetime.date(2025, 9, 5),
        "FinishedDate"  : None,
        "Quantity"      : 1.0,
        "Material"      : "Test vật liệu",
        "Size"          : None,
        "Finish"        : None,
        "Structure"     : "Test kết cấu",
    }

    # Gộp defaults + excel + fields tự sinh
    row = {**DEFAULTS, **from_excel}
    row["RowId"]      = gen_row_id()
    row["CreatedAt"]  = now
    row["ModifiedAt"] = now

    # ── Build câu INSERT ────────────────────────────────────────────────────
    cols = list(row.keys())
    sql  = (
        f"INSERT INTO B20BOM ({', '.join(f'[{c}]' for c in cols)}) "
        f"OUTPUT INSERTED.Id "
        f"VALUES ({', '.join(['?']*len(cols))})"
    )
    values = [row[c] for c in cols]

    print("INSERT B20BOM...")
    print(f"  Code    : {row['Code']}")
    print(f"  DocDate : {row['DocDate']}")
    print(f"  RowId   : {row['RowId']}")
    print(f"  ItemId0 : {row['ItemId0']}")
    print()

    cur = conn.cursor()
    try:
        cur.execute(sql, values)
        new_id = cur.fetchone()[0]
        conn.commit()
        print(f"✅  INSERT thành công — B20BOM.Id = {new_id}")
    except Exception as e:
        conn.rollback()
        print(f"❌  LỖI:\n{e}")
        sys.exit(1)

    # ── Đọc lại để xác nhận ─────────────────────────────────────────────────
    cur.execute("SELECT Id, Code, DocDate, BranchCode, BOMType, IsActive, Version FROM B20BOM WHERE Id = ?", new_id)
    r = cur.fetchone()
    print()
    print("Dữ liệu đã lưu:")
    print(f"  Id={r[0]}  Code={r[1]}  DocDate={r[2]}  BranchCode={r[3]}  BOMType={r[4]}  IsActive={r[5]}  Version={r[6]}")

    conn.close()
    print("\nDone.")

if __name__ == "__main__":
    main()

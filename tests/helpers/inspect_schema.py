"""
tests/helpers/inspect_schema.py
────────────────────────────────
Chạy script này 1 lần để biết chính xác cần điền gì vào test_db_integration.py.

Yêu cầu: VPN bật (cần kết nối DB để lấy schema B20BOM).

Cách chạy (từ thư mục gốc project):
  python tests/helpers/inspect_schema.py

Output:
  1. Columns của HEADER mapping (field Excel → SQL column)
  2. NOT NULL columns của B20BOM và B20BOMDetail
  3. Stored Procedures bắt đầu bằng usp_BOMTool
  4. Copy-paste snippet sẵn để điền vào test_db_integration.py
"""
import sys, os, json, textwrap
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "Tools"))

DB_CFG = os.path.join(os.path.dirname(__file__), "..", "..", "Tools", "config", "db_config.json")


def _divider(title):
    print(f"\n{'─' * 60}")
    print(f"  {title}")
    print('─' * 60)


def inspect_mapping():
    """Đọc HEADER section từ CK_Mapping_v5.xlsx."""
    _divider("HEADER mapping (Excel meta → SQL column B20BOM)")
    try:
        from services.mapping_loader import load_mapping
        mapping = load_mapping()
        header_recs = mapping.get('HEADER', [])
        if not header_recs:
            print("  [!] Không có HEADER section trong mapping")
            return []
        rows = []
        for r in header_recs:
            sql_col   = r.get('sql_col', '')
            nguon     = r.get('nguon_dl', '')
            ten_excel = r.get('ten_excel', '')
            mac_dinh  = r.get('mac_dinh', '')
            kieu_dl   = r.get('kieu_dl', '')
            req       = r.get('bat_buoc', '')
            rows.append((sql_col, nguon, ten_excel, mac_dinh, kieu_dl, req))
            src_info = (
                f"Excel['{ten_excel}']" if ten_excel and nguon == 'Excel'
                else f"nguon={nguon}" + (f", mac_dinh={mac_dinh}" if mac_dinh else "")
            )
            required_mark = "  ← BẮT BUỘC" if str(req) in ('1','True','true') else ""
            print(f"  {sql_col:<30} {kieu_dl:<15} {src_info}{required_mark}")
        return rows
    except Exception as e:
        print(f"  [!] Lỗi đọc mapping: {e}")
        return []


def inspect_db():
    """Kết nối DB, lấy NOT NULL columns của B20BOM và B20BOMDetail."""
    try:
        import pyodbc
        with open(DB_CFG) as f:
            cfg = json.load(f)
        conn_str = (
            f"DRIVER={{{cfg['driver']}}};"
            f"SERVER={cfg['server']};"
            f"DATABASE={cfg['database']};"
            f"UID={cfg['username']};PWD={cfg['password']};"
            "TrustServerCertificate=yes;"
        )
        conn = pyodbc.connect(conn_str, timeout=int(cfg.get('timeout', 10)))
        return conn
    except Exception as e:
        print(f"\n  [!] Không kết nối được DB: {e}")
        print("  Kiểm tra VPN đã bật chưa.")
        return None


def get_not_null_cols(conn, table_name):
    cur = conn.cursor()
    cur.execute("""
        SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = ?
          AND IS_NULLABLE = 'NO'
          AND COLUMN_DEFAULT IS NULL
          AND COLUMNPROPERTY(OBJECT_ID(TABLE_NAME), COLUMN_NAME, 'IsIdentity') = 0
        ORDER BY ORDINAL_POSITION
    """, (table_name,))
    return [(r.COLUMN_NAME, r.DATA_TYPE, r.CHARACTER_MAXIMUM_LENGTH)
            for r in cur.fetchall()]


def get_all_cols(conn, table_name):
    cur = conn.cursor()
    cur.execute("""
        SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE,
               COLUMNPROPERTY(OBJECT_ID(TABLE_NAME), COLUMN_NAME, 'IsIdentity') AS is_identity
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = ?
        ORDER BY ORDINAL_POSITION
    """, (table_name,))
    return [(r.COLUMN_NAME, r.DATA_TYPE, r.IS_NULLABLE, bool(r.is_identity))
            for r in cur.fetchall()]


def get_sps(conn):
    cur = conn.cursor()
    cur.execute("""
        SELECT name FROM sys.objects
        WHERE type = 'P' AND name LIKE 'usp_BOMTool%'
        ORDER BY name
    """)
    return [r.name for r in cur.fetchall()]


def generate_snippet(bom_cols, header_map_rows):
    """Sinh code snippet sẵn để paste vào test_db_integration.py."""
    _divider("SNIPPET — paste vào test_db_integration.py")

    # Filter ra những cột NOT NULL, không phải identity
    not_null_non_id = [(c, dt) for c, dt, nullable, is_id in bom_cols
                       if nullable == 'NO' and not is_id]

    # Map sql_col → nguon từ mapping để gợi ý giá trị
    col_to_src = {}
    if header_map_rows:
        for sql_col, nguon, ten_excel, mac_dinh, kieu_dl, req in header_map_rows:
            col_to_src[sql_col] = (nguon, ten_excel, mac_dinh, kieu_dl)

    lines = ["test_row = {"]
    for col, dt in not_null_non_id:
        src = col_to_src.get(col)
        if src:
            nguon, ten_excel, mac_dinh, kieu_dl = src
            if nguon == 'CoDinh' and mac_dinh:
                hint = f"# mac_dinh={mac_dinh!r} ({nguon})"
                try:    val = repr(int(mac_dinh))
                except: val = repr(mac_dinh)
            elif nguon == 'Excel' and ten_excel:
                val  = repr(f"<từ Excel field '{ten_excel}'>")
                hint = f"# nguon=Excel, ten_excel='{ten_excel}'"
            elif nguon in ('SP', 'TinhToan', 'HeThong'):
                val  = repr(None)
                hint = f"# nguon={nguon} — có thể để NULL hoặc bỏ qua"
            else:
                val  = repr(f"<{dt}>")
                hint = f"# {nguon}"
        else:
            val  = repr(f"<{dt}>")
            hint = "# không có trong mapping"
        lines.append(f'    "{col}": {val},  {hint}')
    lines.append("}")
    print(textwrap.indent("\n".join(lines), "  "))


def main():
    print("\n╔══════════════════════════════════════════════════╗")
    print("║  BOHO Import Tool — DB Schema Inspector         ║")
    print("╚══════════════════════════════════════════════════╝")

    header_map_rows = inspect_mapping()

    conn = inspect_db()
    if not conn:
        print("\nChạy lại sau khi bật VPN để xem schema DB.")
        return

    _divider("B20BOM — tất cả columns (NOT NULL, non-identity cần điền)")
    bom_all = get_all_cols(conn, 'B20BOM')
    for col, dt, nullable, is_id in bom_all:
        mark = []
        if is_id:       mark.append("IDENTITY")
        if nullable == 'NO' and not is_id: mark.append("NOT NULL ← cần điền")
        mark_str = f"  [{', '.join(mark)}]" if mark else ""
        print(f"  {col:<35} {dt:<15}{mark_str}")

    _divider("B20BOMDetail — NOT NULL, non-identity columns")
    detail_nn = get_not_null_cols(conn, 'B20BOMDetail')
    for col, dt, maxlen in detail_nn:
        print(f"  {col:<35} {dt:<15} (maxlen={maxlen})")

    _divider("Stored Procedures usp_BOMTool*")
    sps = get_sps(conn)
    if sps:
        for sp in sps:
            print(f"  {sp}")
    else:
        print("  (không tìm thấy SP nào bắt đầu bằng usp_BOMTool)")

    generate_snippet(bom_all, header_map_rows)

    conn.close()
    print("\n✅  Done. Copy snippet ở trên vào test_db_integration.py.")
    print("   Thay các giá trị '<...>' bằng giá trị thực từ file Excel của bạn.\n")


if __name__ == "__main__":
    main()

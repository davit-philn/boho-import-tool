"""
tests/test_db_integration.py — Option B: DB Integration Tests.

Yêu cầu: VPN bật + SQL Server 172.16.30.20 reachable.
Mỗi test chạy trong TRANSACTION riêng → luôn ROLLBACK sau khi xong.
DB sẽ không bị dirty dù test pass hay fail.

Chạy:
  pytest tests/test_db_integration.py -m integration -v

Cấu trúc:
  - TestDBConnectivity   : smoke test kết nối + schema tồn tại
  - TestBOMHeaderInsert  : INSERT B20BOM header + verify + rollback
  - TestFKConstraints    : verify Foreign Key constraints hoạt động đúng
  - TestStoredProcedures : các SP cần thiết tồn tại và callable
  - TestFullPipeline     : parse Excel → validate → insert header + detail → rollback
"""
import os
import datetime
import pytest

pytestmark = pytest.mark.integration

FIXTURES_DIR = os.path.join(os.path.dirname(__file__), "fixtures")

BOM_TABLE    = "B20BOM"
DETAIL_TABLE = "B20BOMDetail"

# ── Schema-aware row builder ──────────────────────────────────────────────────
#
# Defaults derived from CK_Mapping_v5.xlsx mac_dinh column (CoDinh + HeThong).
# _auto_row() reads the actual NOT NULL columns from INFORMATION_SCHEMA and
# fills them in, so tests don't need to be maintained when the schema changes.

_BOM_KNOWN_DEFAULTS: dict = {
    "BranchCode": "A01",
    "DocCode":    "BOM",
    "DocStatus":  4,
    "BOMType":    "SX",
    "TypeB":      "L2",
    "ModifiedBy": -1,
    "ParentId":   -1,
    "IsGroup":    0,
    "IsActive":   1,
    "EmployeeId": 1,
    "IsDraftData": 0,
    "Version":    "1",
    "CreatedBy":  1,   # FK User — Id=1 (admin) thường tồn tại
}

_SQL_TYPE_SAFE: dict = {
    "varchar": "", "nvarchar": "", "char": "", "nchar": "",
    "int": 0, "bigint": 0, "smallint": 0, "tinyint": 0,
    "bit": 0,
    "numeric": 0, "decimal": 0, "float": 0, "real": 0, "money": 0,
}


def _auto_row(cur, table: str, extra: dict | None = None) -> dict:
    """
    Build minimal INSERT row from schema NOT NULL columns + mapping defaults.
    Skips identity columns. Falls back to zero/empty for unknown types.
    """
    _now   = datetime.datetime.now()
    _today = _now.date()

    cur.execute(
        """
        SELECT COLUMN_NAME, DATA_TYPE
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = ?
          AND IS_NULLABLE = 'NO'
          AND COLUMNPROPERTY(OBJECT_ID(TABLE_SCHEMA + '.' + TABLE_NAME),
                             COLUMN_NAME, 'IsIdentity') = 0
        ORDER BY ORDINAL_POSITION
        """,
        (table,),
    )

    row: dict = {}
    for col_name, data_type in cur.fetchall():
        if col_name in _BOM_KNOWN_DEFAULTS:
            row[col_name] = _BOM_KNOWN_DEFAULTS[col_name]
        elif data_type == "date":
            row[col_name] = _today
        elif data_type in ("datetime", "datetime2", "smalldatetime"):
            row[col_name] = _now
        elif "Id" in col_name or "By" in col_name:
            row[col_name] = 1  # FK placeholder; fails clearly if FK row missing
        elif data_type in _SQL_TYPE_SAFE:
            row[col_name] = _SQL_TYPE_SAFE[data_type]
        # Unknown types omitted → INSERT may fail with a clear SQL error

    if extra:
        row.update(extra)
    return row


# ── 1. Smoke test kết nối và schema ──────────────────────────────────────────

class TestDBConnectivity:
    """Kiểm tra DB accessible và schema đúng trước khi chạy test thực tế."""

    def test_connection_alive(self, db_tx):
        cur = db_tx.cursor()
        cur.execute("SELECT 1 AS ping")
        row = cur.fetchone()
        assert row[0] == 1

    def test_bom_table_exists(self, db_tx):
        cur = db_tx.cursor()
        cur.execute(
            "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES "
            "WHERE TABLE_NAME = ?", (BOM_TABLE,)
        )
        count = cur.fetchone()[0]
        assert count > 0, f"Bảng '{BOM_TABLE}' không tồn tại trong DB"

    def test_detail_table_exists(self, db_tx):
        cur = db_tx.cursor()
        cur.execute(
            "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES "
            "WHERE TABLE_NAME = ?", (DETAIL_TABLE,)
        )
        count = cur.fetchone()[0]
        assert count > 0, f"Bảng '{DETAIL_TABLE}' không tồn tại trong DB"

    def test_bom_table_has_expected_columns(self, db_tx):
        expected_cols = {"Id", "ItemId", "BranchCode"}
        cur = db_tx.cursor()
        cur.execute(
            "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
            "WHERE TABLE_NAME = ?", (BOM_TABLE,)
        )
        actual_cols = {row[0] for row in cur.fetchall()}
        missing = {c for c in expected_cols if c not in actual_cols}
        assert not missing, (
            f"B20BOM thiếu các column: {missing}\n"
            f"Columns hiện có: {sorted(actual_cols)}"
        )


# ── 2. INSERT B20BOM Header + ROLLBACK ───────────────────────────────────────

class TestBOMHeaderInsert:
    """
    Kiểm tra INSERT B20BOM header row.
    Sau mỗi test: db_tx fixture tự động ROLLBACK → không có gì tồn tại trong DB.
    """

    def _build_insert_sql(self, cols: list[str]) -> str:
        return (
            f"INSERT INTO {BOM_TABLE} ({', '.join(f'[{c}]' for c in cols)}) "
            f"OUTPUT INSERTED.Id VALUES ({', '.join(['?'] * len(cols))})"
        )

    def test_insert_valid_bom_header(self, db_tx):
        """
        INSERT 1 BOM header hợp lệ → verify Id trả về dương.
        _auto_row() tự đọc NOT NULL columns từ INFORMATION_SCHEMA và điền
        mapping defaults (CoDinh/HeThong). Nếu FK columns không tồn tại
        với Id=1, test sẽ báo lỗi FK rõ ràng — không skip.
        """
        cur = db_tx.cursor()
        test_row = _auto_row(cur, BOM_TABLE)

        if not test_row:
            pytest.skip("Không đọc được schema B20BOM — kiểm tra kết nối DB")

        cols = list(test_row.keys())
        vals = [test_row[c] for c in cols]

        cur.execute(self._build_insert_sql(cols), vals)
        new_id = cur.fetchone()[0]

        assert new_id is not None and int(new_id) > 0, (
            f"INSERT B20BOM phải trả về Id dương, got: {new_id}"
        )

    def test_insert_then_select_back(self, db_tx):
        """INSERT header → SELECT lại trong cùng transaction → verify data khớp."""
        cur = db_tx.cursor()
        test_row = _auto_row(cur, BOM_TABLE)

        if not test_row:
            pytest.skip("Không đọc được schema B20BOM")

        cols = list(test_row.keys())
        vals = [test_row[c] for c in cols]

        cur.execute(self._build_insert_sql(cols), vals)
        new_id = cur.fetchone()[0]

        cur.execute(f"SELECT Id FROM {BOM_TABLE} WHERE Id = ?", (new_id,))
        found = cur.fetchone()
        assert found is not None, (
            f"SELECT sau INSERT không tìm thấy Id={new_id} — "
            "có thể OUTPUT INSERTED không hoạt động đúng"
        )

    def test_rollback_leaves_no_trace(self, db_conn):
        """
        Verify rằng sau ROLLBACK row không còn tồn tại.
        Dùng db_conn (session-scope) để đọc state sau khi TX đã rollback.
        """
        cur = db_conn.cursor()
        cur.execute(f"SELECT ISNULL(MAX(Id), 0) FROM {BOM_TABLE}")
        max_id_before = cur.fetchone()[0]

        test_row = _auto_row(cur, BOM_TABLE)
        if not test_row:
            pytest.skip("Không đọc được schema B20BOM")

        cols = list(test_row.keys())
        vals = [test_row[c] for c in cols]
        sql = (
            f"INSERT INTO {BOM_TABLE} ({', '.join(f'[{c}]' for c in cols)}) "
            f"OUTPUT INSERTED.Id VALUES ({', '.join(['?'] * len(cols))})"
        )

        db_conn.autocommit = False
        try:
            cur.execute(sql, vals)
            inserted_id = cur.fetchone()[0]
            db_conn.rollback()
        finally:
            db_conn.autocommit = True

        cur.execute(f"SELECT COUNT(*) FROM {BOM_TABLE} WHERE Id = ?", (inserted_id,))
        remaining = cur.fetchone()[0]
        assert remaining == 0, (
            f"Row Id={inserted_id} vẫn còn sau ROLLBACK — isolation bị lỗi!"
        )


# ── 3. Foreign Key Constraints ───────────────────────────────────────────────

class TestFKConstraints:
    """Verify DB enforce FK — insert detail với BOM Id không tồn tại phải fail."""

    def test_detail_without_valid_bom_id_fails(self, db_tx):
        """
        INSERT vào B20BOMDetail với BomId fake (không tồn tại) → phải raise FK violation.
        _auto_row() tự fill NOT NULL columns; fake_bom_id ghi đè FK column.
        """
        fake_bom_id = 999_999_999

        cur = db_tx.cursor()
        # Tìm tên FK column trỏ về B20BOM trong B20BOMDetail
        cur.execute(
            """
            SELECT fk_col.COLUMN_NAME
            FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS rc
            JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE fk_col
                ON fk_col.CONSTRAINT_NAME = rc.CONSTRAINT_NAME
            JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE pk_col
                ON pk_col.CONSTRAINT_NAME = rc.UNIQUE_CONSTRAINT_NAME
            WHERE fk_col.TABLE_NAME = ? AND pk_col.TABLE_NAME = ?
            """,
            (DETAIL_TABLE, BOM_TABLE),
        )
        fk_rows = cur.fetchall()

        if not fk_rows:
            pytest.skip(
                f"Không tìm thấy FK từ {DETAIL_TABLE} → {BOM_TABLE}. "
                "Kiểm tra INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS."
            )

        fk_col = fk_rows[0][0]
        test_detail = _auto_row(cur, DETAIL_TABLE, extra={fk_col: fake_bom_id})

        if not test_detail:
            pytest.skip("Không đọc được schema B20BOMDetail")

        cols = list(test_detail.keys())
        vals = [test_detail[c] for c in cols]
        sql = (
            f"INSERT INTO {DETAIL_TABLE} ({', '.join(f'[{c}]' for c in cols)}) "
            f"VALUES ({', '.join(['?'] * len(cols))})"
        )

        with pytest.raises(Exception) as exc_info:
            cur.execute(sql, vals)
        err = str(exc_info.value)
        assert any(kw in err.upper() for kw in ("FOREIGN KEY", "REFERENCE", "FK")), (
            f"Exception phải là FK violation, nhưng got: {err}"
        )


# ── 4. Stored Procedures tồn tại ─────────────────────────────────────────────

class TestStoredProcedures:
    """
    Verify các SP được tool gọi tồn tại trong DB (hoặc DB được tham chiếu).
    Tên SP lấy từ CK_Mapping_v5.xlsx sheet _SP_CONFIG và SP_HOOK.
    Cross-DB SPs (B10_Boho.dbo.*) được kiểm tra qua [B10_Boho].sys.objects.
    """

    # Extracted từ CK_Mapping_v5.xlsx _SP_CONFIG + SP_HOOK (script mapping_loader)
    REQUIRED_SPS = [
        "B10_Boho.dbo.usp_sys_CreateSttBySeq",
        "B10_Boho.dbo.usp_B20BOM_AutoVersion",
        "usp_sys_AutoNewCode",
        "B10_Boho.dbo.usp_B30BizDoc_DefaultDocNo",
        "B10_Boho.dbo.usp_B30BizDocDemand_AutoVersion",
        "B10_Boho.dbo.usp_B20BOM_Create_ItemCode",
    ]

    def test_required_sps_exist(self, db_tx):
        cur = db_tx.cursor()
        missing = []
        for sp_full in self.REQUIRED_SPS:
            parts = sp_full.split(".")
            sp_name = parts[-1]
            # Cross-DB: "DB.dbo.name" → query [DB].sys.objects
            if len(parts) >= 3:
                db_prefix = parts[0]
                query = (
                    f"SELECT COUNT(*) FROM [{db_prefix}].sys.objects "
                    "WHERE type = 'P' AND name = ?"
                )
            else:
                query = "SELECT COUNT(*) FROM sys.objects WHERE type = 'P' AND name = ?"
            try:
                cur.execute(query, (sp_name,))
                count = cur.fetchone()[0]
                if count == 0:
                    missing.append(sp_full)
            except Exception as exc:
                missing.append(f"{sp_full} [query error: {exc}]")

        assert not missing, (
            "Các SP sau không tồn tại hoặc không truy cập được:\n"
            + "\n".join(f"  - {s}" for s in missing)
        )


# ── 5. Full Pipeline Test ─────────────────────────────────────────────────────

class TestFullPipeline:
    """
    End-to-end: parse Excel → validate → build INSERT SQL → execute → verify → rollback.
    Dùng file fixtures/bom_valid.xlsx (skip nếu file chưa có).
    Toàn bộ chạy trong 1 TRANSACTION rồi ROLLBACK.
    """

    def test_parse_validate_passes(self, good_excel_path, mapping):
        """Step 1: file fixture phải pass validate layer 1."""
        from services.bom_parser import parse_bom_file
        from services.validators import validate_layer1, count_errors
        tables, meta, _, _ = parse_bom_file(good_excel_path)
        val_errors = validate_layer1(tables, meta, mapping)
        n_err, _ = count_errors(val_errors)
        assert n_err == 0, (
            f"File fixture có {n_err} validation errors — "
            "pipeline test không thể tiếp tục với file lỗi."
        )

    def test_bom_id_generated_is_positive(self, db_tx, good_excel_path, mapping):
        """
        Step 2: INSERT B20BOM header với CoDinh/HeThong defaults → Id phải dương.
        Dùng _auto_row() schema-aware thay vì hardcode columns.
        """
        cur = db_tx.cursor()
        bom_row = _auto_row(cur, BOM_TABLE)

        if not bom_row:
            pytest.skip("Không đọc được schema B20BOM")

        cols = list(bom_row.keys())
        vals = [bom_row[c] for c in cols]
        sql = (
            f"INSERT INTO {BOM_TABLE} ({', '.join(f'[{c}]' for c in cols)}) "
            f"OUTPUT INSERTED.Id VALUES ({', '.join(['?'] * len(cols))})"
        )
        cur.execute(sql, vals)
        new_id = cur.fetchone()[0]
        assert int(new_id) > 0, f"BOM Id phải dương, got: {new_id}"

    def test_detail_count_matches_parsed_rows(
        self, db_tx, good_excel_path, mapping, db_config
    ):
        """
        Step 3 (advanced): INSERT header + detail rows → đếm trong DB.
        Đây là integration test phức tạp nhất — cần full pipeline hoạt động.
        """
        pytest.skip(
            "TestFullPipeline.test_detail_count_matches_parsed_rows cần full pipeline.\n"
            "Implement sau khi TestBOMHeaderInsert đã pass."
        )

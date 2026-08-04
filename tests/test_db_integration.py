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

# Đánh dấu toàn bộ module cần integration marker
pytestmark = pytest.mark.integration

FIXTURES_DIR = os.path.join(os.path.dirname(__file__), "fixtures")

# ── Constants ─────────────────────────────────────────────────────────────────
BOM_TABLE    = "B20BOM"
DETAIL_TABLE = "B20BOMDetail"


# ── 1. Smoke test kết nối và schema ──────────────────────────────────────────

class TestDBConnectivity:
    """Kiểm tra DB accessible và schema đúng trước khi chạy test thực tế."""

    def test_connection_alive(self, db_tx):
        """Thực hiện SELECT 1 — pass = DB reachable."""
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

    def test_bom_table_has_expected_columns(self, db_tx, db_config):
        """
        Verify các cột cốt lõi của B20BOM tồn tại.

        ── ĐIỀU CHỈNH DANH SÁCH COLUMN NẾU SCHEMA THỰC TẾ KHÁC ────────────────
        """
        expected_cols = {
            # Tên column (case-insensitive) mà B20BOM phải có
            "Id", "ItemId", "BranchCode",
            # << THÊM COLUMN THỰC TẾ CỦA BẠN VÀO ĐÂY
        }
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
        """Build câu INSERT INTO B20BOM ... OUTPUT INSERTED.Id."""
        return (
            f"INSERT INTO {BOM_TABLE} ({', '.join(f'[{c}]' for c in cols)}) "
            f"OUTPUT INSERTED.Id VALUES ({', '.join(['?'] * len(cols))})"
        )

    def test_insert_valid_bom_header(self, db_tx):
        """
        INSERT 1 BOM header hợp lệ → verify Id trả về dương.

        ── ĐIỀN GIÁ TRỊ BOM HEADER THỰC TẾ CỦA BẠN VÀO ĐÂY ─────────────────
        Col + value phải match schema của bảng B20BOM.
        Lấy từ 1 lần import thực tế (xem Export SQL trong app).
        """
        test_row = {
            # "ItemId":     "SP-TEST-001",    # << ĐIỀN
            # "BranchCode": "TEST",           # << ĐIỀN
            # "Status":     0,                # << ĐIỀN
            # ... các cột NOT NULL khác ...   # << ĐIỀN
        }
        if not test_row:
            pytest.skip(
                "test_row chưa được điền — mở test_db_integration.py "
                "và điền giá trị vào test_insert_valid_bom_header()"
            )

        cols = list(test_row.keys())
        vals = [test_row[c] for c in cols]
        sql  = self._build_insert_sql(cols)

        cur = db_tx.cursor()
        cur.execute(sql, vals)
        new_id = cur.fetchone()[0]

        assert new_id is not None and int(new_id) > 0, (
            f"INSERT B20BOM phải trả về Id dương, got: {new_id}"
        )

    def test_insert_then_select_back(self, db_tx):
        """
        INSERT header → SELECT lại trong cùng transaction → verify data khớp.

        ── ĐIỀN GIÁ TRỊ BOM HEADER THỰC TẾ CỦA BẠN VÀO ĐÂY ─────────────────
        """
        test_row = {
            # "ItemId":     "SP-TEST-SELECT",  # << ĐIỀN
            # "BranchCode": "TEST",            # << ĐIỀN
        }
        if not test_row:
            pytest.skip("test_row chưa được điền")

        cols = list(test_row.keys())
        vals = [test_row[c] for c in cols]

        cur = db_tx.cursor()
        cur.execute(self._build_insert_sql(cols), vals)
        new_id = cur.fetchone()[0]

        # SELECT lại trong cùng TX (chưa commit)
        cur.execute(f"SELECT Id FROM {BOM_TABLE} WHERE Id = ?", (new_id,))
        found = cur.fetchone()
        assert found is not None, (
            f"SELECT sau INSERT không tìm thấy Id={new_id} — "
            "có thể OUTPUT INSERTED không hoạt động đúng"
        )

    def test_rollback_leaves_no_trace(self, db_conn):
        """
        Verify rằng sau ROLLBACK (do db_tx fixture), row không còn tồn tại.
        Test này dùng db_conn (session-scope) chứ không dùng db_tx để
        đọc state SAU KHI transaction ở test trước đã rollback.
        """
        # Lấy Max Id hiện tại trước khi bắt đầu
        cur = db_conn.cursor()
        cur.execute(f"SELECT ISNULL(MAX(Id), 0) FROM {BOM_TABLE}")
        max_id_before = cur.fetchone()[0]

        # Làm 1 insert trong TX rồi tự rollback
        db_conn.autocommit = False
        try:
            test_row = {}  # << ĐIỀN NẾU CÓ — xem test_insert_valid_bom_header()
            if not test_row:
                pytest.skip("test_row chưa được điền")
            cols = list(test_row.keys())
            vals = [test_row[c] for c in cols]
            sql  = (
                f"INSERT INTO {BOM_TABLE} ({', '.join(f'[{c}]' for c in cols)}) "
                f"OUTPUT INSERTED.Id VALUES ({', '.join(['?'] * len(cols))})"
            )
            cur.execute(sql, vals)
            inserted_id = cur.fetchone()[0]
            db_conn.rollback()
        finally:
            db_conn.autocommit = True

        # Verify row đã biến mất
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
        INSERT vào B20BOMDetail với BomId = 999999999 (không tồn tại)
        phải raise exception (FK violation).

        ── ĐIỀN TÊN COLUMN FK VÀ CÁC COLUMN NOT NULL CỦA B20BOMDetail ────────
        """
        fake_bom_id = 999_999_999
        test_detail = {
            # "BomId":          fake_bom_id,   # << ĐIỀN (tên column FK thực tế)
            # "BOMDetailType":  1,             # << ĐIỀN
            # ... các NOT NULL column khác ... # << ĐIỀN
        }
        if not test_detail:
            pytest.skip("test_detail chưa được điền")

        cols = list(test_detail.keys())
        vals = [test_detail[c] for c in cols]
        sql  = (
            f"INSERT INTO {DETAIL_TABLE} ({', '.join(f'[{c}]' for c in cols)}) "
            f"VALUES ({', '.join(['?'] * len(cols))})"
        )
        cur = db_tx.cursor()
        with pytest.raises(Exception) as exc_info:
            cur.execute(sql, vals)
        err = str(exc_info.value)
        # FK violation thường chứa "FOREIGN KEY" hoặc "REFERENCE"
        assert any(kw in err.upper() for kw in ("FOREIGN KEY", "REFERENCE", "FK")), (
            f"Exception phải là FK violation, nhưng got: {err}"
        )


# ── 4. Stored Procedures tồn tại ─────────────────────────────────────────────

class TestStoredProcedures:
    """
    Verify các SP được dùng bởi tool tồn tại trong DB.

    ── ĐIỀN DANH SÁCH SP THỰC TẾ CỦA BẠN VÀO ĐÂY ───────────────────────────
    Lấy từ mapping file: cột SPName trong sheet hooks.
    """
    REQUIRED_SPS = [
        # "usp_BOMTool_DeleteBOM",    # << THÊM VÀO
        # "usp_BOMTool_GetNextBomId", # << THÊM VÀO
        # ...
    ]

    def test_required_sps_exist(self, db_tx):
        if not self.REQUIRED_SPS:
            pytest.skip("REQUIRED_SPS chưa được điền")
        cur = db_tx.cursor()
        for sp_name in self.REQUIRED_SPS:
            cur.execute(
                "SELECT COUNT(*) FROM sys.objects "
                "WHERE type = 'P' AND name = ?", (sp_name,)
            )
            count = cur.fetchone()[0]
            assert count > 0, (
                f"Stored Procedure '{sp_name}' không tồn tại trong DB.\n"
                "Kiểm tra lại tên SP hoặc quyền truy cập."
            )


# ── 5. Full Pipeline Test ─────────────────────────────────────────────────────

class TestFullPipeline:
    """
    End-to-end: parse Excel → validate → build INSERT SQL → execute → verify → rollback.

    Dùng file fixtures/bom_valid.xlsx.
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
        Step 2: INSERT B20BOM header (dùng giá trị từ meta của Excel fixture)
        → Id sinh ra phải dương.

        ── ĐIỀN MAPPING FIELD META → SQL COLUMN CỦA BẠN VÀO ĐÂY ──────────────
        Xem trong main_window.py: phần resolve row để biết meta field nào
        map vào sql column nào của B20BOM.
        """
        from services.bom_parser import parse_bom_file
        _, meta, _, _ = parse_bom_file(good_excel_path)

        # Build row từ meta — điều chỉnh mapping theo schema thực tế
        bom_row = {
            # "ItemId":     meta.get("Mã sản phẩm"),  # << ĐIỀN
            # "BranchCode": meta.get("Dự án"),        # << ĐIỀN
        }
        bom_row = {k: v for k, v in bom_row.items() if v is not None}

        if not bom_row:
            pytest.skip(
                "bom_row chưa được điền — "
                "mở test_db_integration.py::TestFullPipeline::test_bom_id_generated_is_positive"
            )

        cols = list(bom_row.keys())
        vals = [bom_row[c] for c in cols]
        sql  = (
            f"INSERT INTO {BOM_TABLE} ({', '.join(f'[{c}]' for c in cols)}) "
            f"OUTPUT INSERTED.Id VALUES ({', '.join(['?'] * len(cols))})"
        )
        cur = db_tx.cursor()
        cur.execute(sql, vals)
        new_id = cur.fetchone()[0]
        assert int(new_id) > 0, f"BOM Id phải dương, got: {new_id}"

    def test_detail_count_matches_parsed_rows(
        self, db_tx, good_excel_path, mapping, db_config
    ):
        """
        Step 3 (advanced): INSERT header + detail rows → đếm trong DB.
        Đây là integration test phức tạp nhất — cần điền đầy đủ column mapping.

        ── STATUS: SKELETON — cần điền thêm trước khi chạy ───────────────────
        """
        pytest.skip(
            "TestFullPipeline.test_detail_count_matches_parsed_rows chưa được điền.\n"
            "Xem comment trong file và điền column mapping + SP info."
        )

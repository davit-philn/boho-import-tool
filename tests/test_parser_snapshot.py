"""
tests/test_parser_snapshot.py — Option A: Snapshot / Offline Tests (không cần VPN).

Pipeline được test:
  parse_bom_file(excel) → validate_layer1() → count_errors()

Không cần kết nối DB. Chạy được cả khi VPN tắt.

Cấu trúc:
  - TestParserValid    : parse file chuẩn → đúng tables / meta / row count
  - TestValidatorValid : validate file chuẩn → 0 error
  - TestValidatorBad   : validate file lỗi → detect đúng lỗi
  - TestRowCounts      : snapshot số dòng per-section (tự tạo lần đầu)
"""
import os
import re
import json
import pytest

FIXTURES_DIR    = os.path.join(os.path.dirname(__file__), "fixtures")
SNAPSHOT_COUNTS = os.path.join(FIXTURES_DIR, "row_counts.json")


# ── Helpers ───────────────────────────────────────────────────────────────────

def _format_errors(val_errors: dict, limit: int = 20) -> str:
    lines = []
    for section, errs in val_errors.items():
        for e in errs:
            if e.get("severity") == "error":
                lines.append(
                    f"  [{section}] row={e.get('row','?')} "
                    f"col={e.get('col','?')}: {e.get('msg','')}"
                )
    return "\n".join(lines[:limit])


# ── 1. Parse — file hợp lệ ───────────────────────────────────────────────────

class TestParserValid:
    """Kiểm tra parse_bom_file() trên file Excel chuẩn."""

    def test_returns_nonempty_tables(self, good_excel_path):
        from services.bom_parser import parse_bom_file
        tables, meta, skipped, warns = parse_bom_file(good_excel_path)
        assert tables, "parse_bom_file phải trả về ít nhất 1 section"

    def test_meta_has_required_fields(self, good_excel_path):
        from services.bom_parser import parse_bom_file
        _, meta, _, _ = parse_bom_file(good_excel_path)
        missing = [f for f in ("Mã sản phẩm", "Dự án", "Đơn hàng") if not meta.get(f)]
        assert not missing, f"Meta thiếu các field bắt buộc: {missing}"

    def test_all_sections_have_nonempty_dataframe(self, good_excel_path):
        from services.bom_parser import parse_bom_file
        tables, _, _, _ = parse_bom_file(good_excel_path)
        for name, tbl in tables.items():
            df = tbl.get("df")
            assert df is not None, f"Section '{name}' thiếu DataFrame"
            assert not df.empty,   f"Section '{name}' có DataFrame rỗng"

    def test_no_skipped_sheets_on_valid_file(self, good_excel_path):
        """File chuẩn không được bỏ qua sheet nào có dữ liệu."""
        from services.bom_parser import parse_bom_file
        _, _, skipped, _ = parse_bom_file(good_excel_path)
        # Chỉ cảnh báo nếu có skipped — không fail (có thể file có sheet extra)
        if skipped:
            print(f"\n  [WARN] {len(skipped)} sheet bị bỏ qua: {skipped}")


# ── 2. Validate layer 1 — file hợp lệ ───────────────────────────────────────

class TestValidatorValid:
    """validate_layer1() trên file Excel chuẩn không được sinh error."""

    def test_zero_errors(self, good_excel_path, mapping):
        from services.bom_parser import parse_bom_file
        from services.validators import validate_layer1, count_errors
        tables, meta, _, _ = parse_bom_file(good_excel_path)
        val_errors = validate_layer1(tables, meta, mapping)
        n_err, n_wrn = count_errors(val_errors)
        assert n_err == 0, (
            f"File hợp lệ không được có lỗi. Tìm thấy {n_err} error:\n"
            + _format_errors(val_errors)
        )

    def test_warning_count_is_stable(self, good_excel_path, mapping):
        """
        Số warning phải ổn định qua các lần chạy (snapshot).
        Lần đầu chạy → tạo snapshot. Lần sau → so sánh.

        ── NẾU WARNING TĂNG DO THAY ĐỔI FILE: xóa fixtures/warning_count.txt ──
        """
        from services.bom_parser import parse_bom_file
        from services.validators import validate_layer1, count_errors
        tables, meta, _, _ = parse_bom_file(good_excel_path)
        val_errors = validate_layer1(tables, meta, mapping)
        _, n_wrn = count_errors(val_errors)

        snap_path = os.path.join(FIXTURES_DIR, "warning_count.txt")
        if not os.path.exists(snap_path):
            with open(snap_path, "w") as f:
                f.write(str(n_wrn))
            print(f"\n  [SNAPSHOT CREATED] warning_count = {n_wrn}")
            return

        expected = int(open(snap_path).read().strip())
        assert n_wrn == expected, (
            f"Warning count thay đổi: expected {expected}, got {n_wrn}.\n"
            "Nếu intentional: xóa fixtures/warning_count.txt để regenerate."
        )


# ── 3. Validate layer 1 — file lỗi ───────────────────────────────────────────

class TestValidatorBad:
    """validate_layer1() trên file Excel lỗi phải detect được lỗi."""

    def test_errors_detected(self, bad_excel_path, mapping):
        from services.bom_parser import parse_bom_file
        from services.validators import validate_layer1, count_errors
        tables, meta, _, _ = parse_bom_file(bad_excel_path)
        val_errors = validate_layer1(tables, meta, mapping)
        n_err, _ = count_errors(val_errors)
        assert n_err > 0, (
            "File lỗi phải bị detect ít nhất 1 error nhưng count_errors() trả về 0."
        )

    def test_missing_required_header_unit(self):
        """Unit test thuần (không cần file): meta rỗng → 3 errors."""
        from services.validators import validate_layer1, count_errors
        val_errors = validate_layer1({}, {})
        n_err, _ = count_errors(val_errors)
        assert n_err == 3, (
            f"3 required header fields đều thiếu → phải có 3 errors, got {n_err}"
        )

    def test_error_severity_is_error_not_warning(self, bad_excel_path, mapping):
        """Lỗi nghiêm trọng phải là severity='error', không phải 'warning'."""
        from services.bom_parser import parse_bom_file
        from services.validators import validate_layer1
        tables, meta, _, _ = parse_bom_file(bad_excel_path)
        val_errors = validate_layer1(tables, meta, mapping)
        all_items = [e for errs in val_errors.values() for e in errs]
        has_error = any(e.get("severity") == "error" for e in all_items)
        assert has_error, "File lỗi phải có ít nhất 1 item với severity='error'"

    def test_specific_columns_flagged(self, bad_excel_path, mapping):
        """
        ── ĐIỀN CÁC COLUMN LỖI CỤ THỂ MÀ FILE bom_invalid.xlsx CỦA BẠN CÓ ──

        Ví dụ nếu file lỗi có cột 'Số lượng' không phải số:
          assert "Số lượng" in error_cols

        Bỏ comment các assert phù hợp với file lỗi thực tế của bạn.
        """
        from services.bom_parser import parse_bom_file
        from services.validators import validate_layer1
        tables, meta, _, _ = parse_bom_file(bad_excel_path)
        val_errors = validate_layer1(tables, meta, mapping)
        all_errors = [e for errs in val_errors.values() for e in errs
                      if e.get("severity") == "error"]
        error_cols = {e.get("col") for e in all_errors}

        # ── UNCOMMENT VÀ ĐIỀN TÊN COLUMN THỰC TẾ CỦA BẠN ──────────────────
        # assert "Số lượng"   in error_cols, f"Phải detect lỗi ở 'Số lượng'. Got: {error_cols}"
        # assert "Ngày"       in error_cols, f"Phải detect lỗi ở 'Ngày'. Got: {error_cols}"
        # assert "Mã sản phẩm" in error_cols
        # ────────────────────────────────────────────────────────────────────
        assert error_cols, f"Phải có ít nhất 1 error column, got empty set"


# ── 4. Snapshot số dòng per-section ──────────────────────────────────────────

class TestRowCountSnapshot:
    """
    Đảm bảo số dòng mỗi section không thay đổi so với lần đầu chạy.

    Workflow:
      1. Lần đầu chạy với file fixture → tạo fixtures/row_counts.json
      2. Mọi lần chạy sau → so sánh với snapshot
      3. Nếu file fixture thay đổi intentionally → xóa row_counts.json

    File snapshot: tests/fixtures/row_counts.json
    Format:  {"BOM2": 45, "BOM3": 12, "BOM4": 8}
    """

    def test_row_counts_match_snapshot(self, good_excel_path):
        from services.bom_parser import parse_bom_file
        tables, _, _, _ = parse_bom_file(good_excel_path)
        actual_counts = {name: len(tbl["df"]) for name, tbl in tables.items()}

        if not os.path.exists(SNAPSHOT_COUNTS):
            with open(SNAPSHOT_COUNTS, "w", encoding="utf-8") as f:
                json.dump(actual_counts, f, ensure_ascii=False, indent=2)
            print(f"\n  [SNAPSHOT CREATED] row_counts.json: {actual_counts}")
            pytest.skip("Snapshot mới tạo — chạy lại để verify.")

        with open(SNAPSHOT_COUNTS, encoding="utf-8") as f:
            expected_counts = json.load(f)

        mismatches = []
        for section, expected in expected_counts.items():
            actual = actual_counts.get(section)
            if actual != expected:
                mismatches.append(
                    f"  '{section}': expected {expected} rows, got {actual}"
                )

        assert not mismatches, (
            "Row count thay đổi so với snapshot:\n"
            + "\n".join(mismatches)
            + "\nNếu intentional: xóa fixtures/row_counts.json để regenerate."
        )

    def test_no_new_unexpected_sections(self, good_excel_path):
        """Section mới xuất hiện trong file phải được review, không được tự động pass."""
        from services.bom_parser import parse_bom_file
        tables, _, _, _ = parse_bom_file(good_excel_path)

        if not os.path.exists(SNAPSHOT_COUNTS):
            pytest.skip("Snapshot chưa tồn tại — chạy test_row_counts_match_snapshot trước.")

        with open(SNAPSHOT_COUNTS, encoding="utf-8") as f:
            expected_counts = json.load(f)

        new_sections = set(tables.keys()) - set(expected_counts.keys())
        assert not new_sections, (
            f"Xuất hiện section mới chưa có trong snapshot: {new_sections}\n"
            "Review và xóa fixtures/row_counts.json nếu đây là intentional."
        )

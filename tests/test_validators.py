"""tests/test_validators.py — Unit tests for services/validators.py."""
import pytest
import pandas as pd
from services.validators import validate_layer1, count_errors, _try_parse_date


# ─── _try_parse_date ─────────────────────────────────────────────────────────

class TestTryParseDate:
    def test_valid_dd_mm_yyyy(self):
        assert _try_parse_date("01/12/2024", "%d/%m/%Y")

    def test_valid_yyyy_mm_dd(self):
        assert _try_parse_date("2024-12-01", "%Y-%m-%d")

    def test_valid_dd_mm_yy(self):
        assert _try_parse_date("01/12/24", "%d/%m/%y")

    def test_invalid_string_returns_false(self):
        assert not _try_parse_date("không phải ngày", "%d/%m/%Y")

    def test_wrong_format_returns_false(self):
        assert not _try_parse_date("2024/01/31", "%d/%m/%Y")

    def test_none_returns_false(self):
        assert not _try_parse_date(None, "%d/%m/%Y")

    def test_invalid_day_31_feb(self):
        assert not _try_parse_date("31/02/2024", "%d/%m/%Y")


# ─── validate_layer1 — header ────────────────────────────────────────────────

class TestValidateLayer1Header:
    def test_all_required_fields_present_no_errors(self):
        meta = {"Mã sản phẩm": "SP001", "Dự án": "DA001", "Đơn hàng": "DH001"}
        result = validate_layer1({}, meta)
        assert result["[H] Phiếu Header"] == []

    def test_missing_one_field_generates_error(self):
        meta = {"Mã sản phẩm": "", "Dự án": "DA001", "Đơn hàng": "DH001"}
        result = validate_layer1({}, meta)
        h_errors = result["[H] Phiếu Header"]
        assert any(e["col"] == "Mã sản phẩm" and e["severity"] == "error"
                   for e in h_errors)

    def test_all_required_fields_missing_three_errors(self):
        result = validate_layer1({}, {})
        h_errors = result["[H] Phiếu Header"]
        assert len(h_errors) == 3

    def test_none_value_passes_as_str_none(self):
        # str(None) = "None" → validator coi là có giá trị; đây là behavior hiện tại
        meta = {"Mã sản phẩm": None, "Dự án": "DA001", "Đơn hàng": "DH001"}
        result = validate_layer1({}, meta)
        h_errors = result["[H] Phiếu Header"]
        # None không bị báo lỗi vì str(None) = "None" (truthy)
        assert not any(e["col"] == "Mã sản phẩm" for e in h_errors)

    def test_whitespace_only_treated_as_missing(self):
        meta = {"Mã sản phẩm": "   ", "Dự án": "DA001", "Đơn hàng": "DH001"}
        result = validate_layer1({}, meta)
        h_errors = result["[H] Phiếu Header"]
        assert any(e["col"] == "Mã sản phẩm" for e in h_errors)


# ─── validate_layer1 — numeric column checking ───────────────────────────────

META_OK = {"Mã sản phẩm": "SP001", "Dự án": "DA001", "Đơn hàng": "DH001"}


def make_table(col_name, values):
    df = pd.DataFrame({col_name: values})
    return {f"[BOM] test": {"df": df, "type": "BOM2"}}


class TestValidateLayer1NumericColumns:
    def test_valid_integer_strings_no_error(self):
        result = validate_layer1(make_table("Số lượng", ["1", "2", "10"]), META_OK)
        assert result["[BOM] test"] == []

    def test_valid_decimal_with_dot_no_error(self):
        result = validate_layer1(make_table("Số lượng", ["1.5", "2.3"]), META_OK)
        assert result["[BOM] test"] == []

    def test_valid_decimal_with_comma_no_error(self):
        result = validate_layer1(make_table("Số lượng", ["1,5"]), META_OK)
        assert result["[BOM] test"] == []

    def test_non_numeric_string_generates_error(self):
        result = validate_layer1(make_table("Số lượng", ["abc"]), META_OK)
        errors = result["[BOM] test"]
        assert any(e["severity"] == "error" for e in errors)

    def test_none_in_numeric_col_skipped(self):
        result = validate_layer1(make_table("Số lượng", [None]), META_OK)
        assert result["[BOM] test"] == []

    def test_nan_float_in_numeric_col_skipped(self):
        result = validate_layer1(make_table("Số lượng", [float("nan")]), META_OK)
        assert result["[BOM] test"] == []

    def test_placeholder_dash_in_numeric_col_skipped(self):
        result = validate_layer1(make_table("Số lượng", ["-", "--", "---"]), META_OK)
        assert result["[BOM] test"] == []

    def test_quantity_english_col_checked(self):
        result = validate_layer1(make_table("Quantity", ["xyz"]), META_OK)
        errors = result["[BOM] test"]
        assert any(e["severity"] == "error" for e in errors)

    def test_empty_detail_table_no_errors(self):
        df = pd.DataFrame()
        tables = {"[BOM] empty": {"df": df, "type": "BOM2"}}
        result = validate_layer1(tables, META_OK)
        assert result["[BOM] empty"] == []


# ─── count_errors ─────────────────────────────────────────────────────────────

class TestCountErrors:
    def test_counts_errors_and_warnings_separately(self):
        val_errors = {
            "t1": [{"severity": "error"}, {"severity": "error"}, {"severity": "warning"}],
            "t2": [{"severity": "warning"}],
        }
        n_err, n_wrn = count_errors(val_errors)
        assert n_err == 2
        assert n_wrn == 2

    def test_empty_dict_returns_zeros(self):
        assert count_errors({}) == (0, 0)

    def test_empty_list_returns_zeros(self):
        assert count_errors({"t": []}) == (0, 0)

    def test_only_errors(self):
        n_err, n_wrn = count_errors({"t": [{"severity": "error"}] * 5})
        assert n_err == 5
        assert n_wrn == 0

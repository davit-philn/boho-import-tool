"""tests/test_utils.py — Unit tests for helper functions in services/utils.py."""
import math
import pytest
from services.utils import _norm_vn, _nan_str


# ─── _norm_vn ────────────────────────────────────────────────────────────────

class TestNormVn:
    def test_removes_diacritics(self):
        assert _norm_vn("tiếng Việt") == "tiengviet"

    def test_d_with_stroke_becomes_d(self):
        assert _norm_vn("đơn hàng") == "donhang"

    def test_uppercase_normalized_to_lowercase(self):
        assert _norm_vn("MÃ SẢN PHẨM") == "masanpham"

    def test_special_chars_removed(self):
        assert _norm_vn("Kính/Thủy (Theo Mẫu)") == "kinhthuytheomau"

    def test_numbers_kept(self):
        assert _norm_vn("BOM2") == "bom2"

    def test_pure_latin_stripped_and_lowercased(self):
        assert _norm_vn("Hello World") == "helloworld"

    def test_empty_string_returns_empty(self):
        assert _norm_vn("") == ""

    def test_only_special_chars_returns_empty(self):
        assert _norm_vn("---///") == ""

    def test_strips_leading_trailing_whitespace(self):
        assert _norm_vn("  kính  ") == "kinh"

    def test_typical_bom_column_name(self):
        assert _norm_vn("Số lượng") == "soluong"

    def test_typical_thdm_column_name(self):
        assert _norm_vn("Định mức") == "dinhmuc"

    def test_combined_vn_latin_numbers(self):
        assert _norm_vn("BOM_I v2") == "bomiv2"


# ─── _nan_str ────────────────────────────────────────────────────────────────

class TestNanStr:
    def test_none_returns_default_empty(self):
        assert _nan_str(None) == ""

    def test_float_nan_returns_default_empty(self):
        assert _nan_str(float("nan")) == ""

    def test_nan_string_lowercase_returns_empty(self):
        assert _nan_str("nan") == ""

    def test_nan_string_mixed_case_returns_empty(self):
        assert _nan_str("NaN") == ""

    def test_empty_string_returns_empty(self):
        assert _nan_str("") == ""

    def test_whitespace_only_returns_empty(self):
        assert _nan_str("   ") == ""

    def test_real_string_returned_stripped(self):
        assert _nan_str("  hello  ") == "hello"

    def test_zero_integer_returned_as_string(self):
        assert _nan_str(0) == "0"

    def test_zero_float_returned_as_string(self):
        assert _nan_str(0.0) == "0.0"

    def test_positive_number_returned(self):
        assert _nan_str(42) == "42"

    def test_custom_default_for_none(self):
        assert _nan_str(None, default="N/A") == "N/A"

    def test_custom_default_for_float_nan(self):
        assert _nan_str(float("nan"), default="N/A") == "N/A"

    def test_real_value_ignores_custom_default(self):
        assert _nan_str("hello", default="N/A") == "hello"

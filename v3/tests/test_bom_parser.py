"""
tests/test_bom_parser.py — Unit tests cho core parsing logic trong bom_parser.py

Covers:
  - NUMERIC_STT_PATTERN / SECTION_STT_PATTERN regex
  - _is_roman_numeral()
  - _parse_section_excel_rows() — 2 bug đã fix:
      Bug 1: Glass items (STT số, ItemId rỗng, ItemName all-caps) không bị bỏ qua
      Bug 2: Mục 'I' với roman_as_hdr=True được treat là section header
"""
import pytest
from services.bom_parser import _is_roman_numeral, _parse_section_excel_rows
from services.mapping_loader import SECTION_STT_PATTERN, NUMERIC_STT_PATTERN


# ─── Helpers ────────────────────────────────────────────────────────────────

# headers: 0=STT, 1=ItemType, 2=ItemId, 3=ItemName
HEADERS = ["STT", "ItemType", "ItemId", "ItemName"]
STT_IDX = 0
COL_MAP = {1: "ItemType", 2: "ItemId", 3: "ItemName"}


def mapping_recs(ff_itemtype: bool = True):
    return [
        {"sql_col": "ItemType", "fill_forward": "1" if ff_itemtype else "0"},
        {"sql_col": "ItemId",   "fill_forward": "0"},
        {"sql_col": "ItemName", "fill_forward": "0"},
    ]


def parse(rows, *, skip_no_item=True, ff_always_override=False,
          roman_as_hdr=False, ff_it=True):
    return _parse_section_excel_rows(
        rows, HEADERS, COL_MAP, mapping_recs(ff_it), STT_IDX,
        skip_no_item=skip_no_item,
        ff_always_override=ff_always_override,
        roman_as_hdr=roman_as_hdr,
    )


# ─── NUMERIC_STT_PATTERN ────────────────────────────────────────────────────

class TestNumericSTTPattern:
    def test_integers(self):
        for v in ("1", "2", "10", "100"):
            assert NUMERIC_STT_PATTERN.match(v), f"'{v}' phải match"

    def test_decimal(self):
        for v in ("1.1", "2.3", "1.1.1"):
            assert NUMERIC_STT_PATTERN.match(v)

    def test_letters_no_match(self):
        for v in ("A", "I", "H", "E1", ""):
            assert not NUMERIC_STT_PATTERN.match(v), f"'{v}' không được match"


# ─── SECTION_STT_PATTERN ────────────────────────────────────────────────────

class TestSectionSTTPattern:
    def test_single_letters(self):
        for ch in "ABCDEFGHIJ":
            assert SECTION_STT_PATTERN.match(ch), f"'{ch}' phải match"

    def test_letter_with_number(self):
        assert SECTION_STT_PATTERN.match("E1")
        assert SECTION_STT_PATTERN.match("E2")

    def test_roman_looking(self):
        # I, V, X cũng là section letter hợp lệ
        for v in ("I", "V", "X"):
            assert SECTION_STT_PATTERN.match(v)

    def test_pure_numeric_no_match(self):
        for v in ("1", "2", "10"):
            assert not SECTION_STT_PATTERN.match(v)


# ─── _is_roman_numeral ──────────────────────────────────────────────────────

class TestIsRomanNumeral:
    def test_valid_roman(self):
        for v in ("I", "II", "III", "IV", "V", "X", "XI", "I."):
            assert _is_roman_numeral(v), f"'{v}' phải là Roman numeral"

    def test_section_letters_not_roman(self):
        for v in ("A", "B", "C", "H", "J"):
            assert not _is_roman_numeral(v), f"'{v}' không phải Roman numeral"

    def test_numbers_not_roman(self):
        for v in ("1", "2", "10"):
            assert not _is_roman_numeral(v)

    def test_empty_not_roman(self):
        assert not _is_roman_numeral("")


# ─── Bug Fix 1: Glass items (KÍNH THỦY THEO MẪU) ───────────────────────────

class TestGlassItemsNotSkipped:
    """
    Bug gốc: STT số + ItemId rỗng + ItemName all-caps → bị skip do rule
    'group header'. Fix: nếu STT là NUMERIC thì giữ lại là data row thực.
    """

    def test_glass_item_numeric_stt_is_kept(self):
        """STT='1', ItemId=None, ItemName='KÍNH THỦY THEO MẪU' → phải có trong output."""
        rows = [("1", None, None, "KÍNH THỦY THEO MẪU")]
        result = parse(rows)
        assert len(result) == 1, "Glass item với numeric STT phải được giữ lại"
        assert result[0]["ItemName"] == "KÍNH THỦY THEO MẪU"

    def test_multiple_glass_items(self):
        rows = [
            ("1", None, None, "KÍNH THỦY THEO MẪU"),
            ("2", None, None, "KÍNH CƯỜNG LỰC"),
        ]
        result = parse(rows)
        assert len(result) == 2

    def test_group_header_all_caps_no_stt_skipped(self):
        """STT=None, ItemId=None, ItemName='VẬT TƯ CHÍNH' → phải bị bỏ qua."""
        rows = [(None, None, None, "VẬT TƯ CHÍNH")]
        result = parse(rows)
        assert len(result) == 0, "Group header all-caps không có STT phải bị bỏ qua"

    def test_normal_row_with_item_id(self):
        rows = [("1", None, "M001", "Ống thép")]
        result = parse(rows)
        assert len(result) == 1
        assert result[0]["ItemId"] == "M001"


# ─── Bug Fix 2: Section 'I' với roman_as_hdr ────────────────────────────────

class TestSectionIRomanAsHdr:
    """
    Bug gốc: Mục 'I BAO BÌ' bị _is_roman_numeral('I')=True → fall-through
    như data row thay vì capture Fill_Forward. Fix: roman_as_hdr=True cho THDM.
    """

    def test_section_I_roman_as_hdr_true_skipped_and_ff_captured(self):
        """roman_as_hdr=True: 'I' là section header → skip + ItemType kế thừa."""
        rows = [
            ("I",  "BAO BÌ", None,    None),   # section header
            ("1",  None,     "BP001", "Hộp carton"),  # data row
        ]
        result = parse(rows, roman_as_hdr=True, ff_always_override=True)
        assert len(result) == 1, "Section 'I' phải bị skip"
        assert result[0]["ItemType"] == "BAO BÌ", "Fill_Forward phải kế thừa từ section I"
        assert result[0]["ItemId"] == "BP001"

    def test_section_I_roman_as_hdr_false_is_data_row(self):
        """roman_as_hdr=False (BOM default): 'I' fall-through như data row."""
        rows = [("I", None, "ROW001", "Some item")]
        result = parse(rows, roman_as_hdr=False)
        assert len(result) == 1, "'I' không bị skip khi roman_as_hdr=False"
        assert result[0]["ItemId"] == "ROW001"

    def test_section_H_always_skipped(self):
        """'H' không phải Roman numeral → luôn là section header, bị skip."""
        rows = [
            ("H", "THẦU PHỤ GIA CÔNG", None,  None),
            ("1", None,                "M001", "Kính thủy"),
        ]
        result = parse(rows, ff_always_override=True)
        assert len(result) == 1
        assert result[0]["ItemType"] == "THẦU PHỤ GIA CÔNG"

    def test_section_J_skipped(self):
        rows = [
            ("J", "SƠN HÓA CHẤT", None,  None),
            ("1", None,           "S001", "Sơn xịt"),
        ]
        result = parse(rows, ff_always_override=True)
        assert len(result) == 1
        assert result[0]["ItemType"] == "SƠN HÓA CHẤT"


# ─── Footer & empty rows ─────────────────────────────────────────────────────

class TestSTTEdgeCases:
    """C4: edge cases cho STT — trailing whitespace, float value từ openpyxl."""

    def test_float_stt_treated_as_numeric(self):
        """openpyxl trả STT = 1.0 (float) thay vì int → phải match NUMERIC_STT_PATTERN."""
        rows = [(1.0, None, "M001", "Ống thép")]
        result = parse(rows)
        assert len(result) == 1, "STT=1.0 (float) phải được giữ lại"

    def test_float_stt_2_treated_as_numeric(self):
        rows = [(2.0, None, "M002", "Kính cường lực")]
        result = parse(rows)
        assert len(result) == 1

    def test_stt_with_trailing_whitespace_kept(self):
        """STT = ' 1 ' (string với khoảng trắng) → strip → '1' → numeric."""
        rows = [(" 1 ", None, "M001", "Ống thép")]
        result = parse(rows)
        assert len(result) == 1, "STT có khoảng trắng phải được giữ lại sau khi strip"

    def test_section_letter_with_whitespace_is_header(self):
        """STT = ' H ' → strip → 'H' → SECTION_STT_PATTERN → section header."""
        rows = [
            (" H ", "THẦU PHỤ", None, None),
            ("1",   None,       "M001", "Ống thép"),
        ]
        result = parse(rows, ff_always_override=True)
        assert len(result) == 1, "'H' sau khi strip là section header → skip"
        assert result[0]["ItemType"] == "THẦU PHỤ"

    def test_stt_zero_int_skipped(self):
        """STT = 0 (int) → bỏ qua (dòng trống trong BOM III)."""
        rows = [(0, None, None, None), ("1", None, "M001", "Item")]
        result = parse(rows)
        assert len(result) == 1

    def test_stt_zero_float_skipped(self):
        """STT = 0.0 (float) → bỏ qua."""
        rows = [(0.0, None, None, None), ("1", None, "M001", "Item")]
        result = parse(rows)
        assert len(result) == 1

    def test_mixed_case_item_name_no_stt_is_kept(self):
        """STT=None, ItemName mixed-case với có data → được giữ lại (không filter all-caps)."""
        rows = [(None, None, None, "Kính Thủy Theo Mẫu")]
        result = parse(rows)
        # ItemName là data key → _is_data_row = True → dòng được giữ lại
        assert len(result) == 1, "Mixed-case ItemName (có data key) phải được giữ lại"


class TestFooterAndEmpty:
    def test_footer_stops_processing(self):
        rows = [
            ("1", None, "M001", "Sản phẩm A"),
            (None, None, None, "Tổng cộng"),
            ("2", None, "M002", "Sản phẩm B"),   # không nên xuất hiện
        ]
        result = parse(rows)
        assert len(result) == 1
        assert result[0]["ItemId"] == "M001"

    def test_empty_rows_skipped(self):
        rows = [
            (None, None, None, None),
            ("1",  None, "M001", "Sản phẩm"),
            (None, None, None, None),
        ]
        result = parse(rows)
        assert len(result) == 1

    def test_all_empty_returns_empty(self):
        rows = [(None, None, None, None)] * 3
        assert parse(rows) == []

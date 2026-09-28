"""tests/test_sheet_detection.py — Unit tests for _detect_sheet_type() in bom_parser.py.

C3 from audit plan: sheet type detection is critical — wrong detection causes entire
file to be parsed under wrong section mapping.
"""
import pytest
from services.bom_parser import _detect_sheet_type


# Synthetic config mirroring a typical CK_Mapping_v5 _CONFIG sheet.
# Uses word-boundary matching: 'BOM I' != 'BOM II' != 'BOM III'
SHEET_CFG = [
    {
        "section": "BOM2", "label": "BOM I",
        "contains": ["BOM", "I"],
        "excludes": ["II", "III", "IV", "V"],
        "start_row": "AUTO",
    },
    {
        "section": "BOM3", "label": "BOM II",
        "contains": ["BOM", "II"],
        "excludes": ["III", "IV"],
        "start_row": "AUTO",
    },
    {
        "section": "BOM4", "label": "BOM III",
        "contains": ["BOM", "III"],
        "excludes": [],
        "start_row": "AUTO",
    },
    {
        "section": "THVT", "label": "TH VT",
        "contains": ["TH", "VT"],
        "excludes": [],
        "start_row": "AUTO",
    },
]


class TestDetectSheetType:
    # ── Basic matches ──────────────────────────────────────────────────────

    def test_bom_i_matches_bom2(self):
        assert _detect_sheet_type("BOM I", sheet_config=SHEET_CFG) == "BOM2"

    def test_bom_ii_matches_bom3(self):
        assert _detect_sheet_type("BOM II", sheet_config=SHEET_CFG) == "BOM3"

    def test_bom_iii_matches_bom4(self):
        assert _detect_sheet_type("BOM III", sheet_config=SHEET_CFG) == "BOM4"

    def test_thvt_matches(self):
        assert _detect_sheet_type("TH VT", sheet_config=SHEET_CFG) == "THVT"

    # ── Separator normalization ────────────────────────────────────────────

    def test_bom_i_underscore_normalized(self):
        assert _detect_sheet_type("BOM_I", sheet_config=SHEET_CFG) == "BOM2"

    def test_thvt_underscore_normalized(self):
        assert _detect_sheet_type("TH_VT", sheet_config=SHEET_CFG) == "THVT"

    def test_bom_ii_underscore_normalized(self):
        assert _detect_sheet_type("BOM_II", sheet_config=SHEET_CFG) == "BOM3"

    # ── Case normalization ─────────────────────────────────────────────────

    def test_lowercase_input_matched(self):
        assert _detect_sheet_type("bom i", sheet_config=SHEET_CFG) == "BOM2"

    def test_mixed_case_input_matched(self):
        assert _detect_sheet_type("Bom I", sheet_config=SHEET_CFG) == "BOM2"

    # ── Word-boundary disambiguation ───────────────────────────────────────

    def test_bom_iii_does_not_match_bom3(self):
        # 'BOM III' contains 'III', should NOT match BOM3 which needs 'II'
        # because \bII\b doesn't match in 'III'
        assert _detect_sheet_type("BOM III", sheet_config=SHEET_CFG) != "BOM3"

    def test_bom_iii_does_not_match_bom2(self):
        # 'BOM III' should NOT match BOM2 — 'I' excluded by 'III' in excludes
        # (actually 'I' does not match via \bI\b in 'III')
        assert _detect_sheet_type("BOM III", sheet_config=SHEET_CFG) != "BOM2"

    def test_bom_ii_does_not_match_bom2(self):
        # 'BOM II' — \bI\b should NOT match inside 'II'
        assert _detect_sheet_type("BOM II", sheet_config=SHEET_CFG) != "BOM2"

    # ── Unknown / fallback ─────────────────────────────────────────────────

    def test_unknown_sheet_name_returns_unknown(self):
        assert _detect_sheet_type("RANDOM SHEET", sheet_config=SHEET_CFG) == "UNKNOWN"

    def test_empty_config_returns_unknown(self):
        assert _detect_sheet_type("BOM I", sheet_config=[]) == "UNKNOWN"

    def test_no_config_kwarg_returns_unknown(self):
        assert _detect_sheet_type("BOM I") == "UNKNOWN"

    def test_empty_sheet_name_returns_unknown(self):
        assert _detect_sheet_type("", sheet_config=SHEET_CFG) == "UNKNOWN"

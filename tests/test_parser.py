"""
Unit tests untuk verifikasi bahwa parser functions sudah dihapus.
Remarks sekarang di-copy langsung ke kolom Request tanpa parsing.
"""

import sys
import os

# Tambahkan root project ke path agar bisa import main
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import main


class TestParserFunctionsRemoved:
    """Verifikasi bahwa semua fungsi parser sudah dihapus dari main.py."""

    def test_parse_remarks_removed(self):
        """Fungsi parse_remarks tidak ada lagi."""
        assert not hasattr(main, 'parse_remarks')

    def test_parse_remarks_mitra_hit_removed(self):
        """Fungsi parse_remarks_mitra_hit tidak ada lagi."""
        assert not hasattr(main, 'parse_remarks_mitra_hit')

    def test_parse_remarks_bss_hit_removed(self):
        """Fungsi parse_remarks_bss_hit tidak ada lagi."""
        assert not hasattr(main, 'parse_remarks_bss_hit')

    def test_parse_with_labels_removed(self):
        """Fungsi _parse_with_labels tidak ada lagi."""
        assert not hasattr(main, '_parse_with_labels')

    def test_parse_by_content_detection_removed(self):
        """Fungsi _parse_by_content_detection tidak ada lagi."""
        assert not hasattr(main, '_parse_by_content_detection')

    def test_parse_bss_single_section_removed(self):
        """Fungsi _parse_bss_single_section tidak ada lagi."""
        assert not hasattr(main, '_parse_bss_single_section')

    def test_parse_bss_multi_section_removed(self):
        """Fungsi _parse_bss_multi_section tidak ada lagi."""
        assert not hasattr(main, '_parse_bss_multi_section')

    def test_extract_url_from_text_removed(self):
        """Fungsi _extract_url_from_text tidak ada lagi."""
        assert not hasattr(main, '_extract_url_from_text')

    def test_extract_json_blocks_removed(self):
        """Fungsi _extract_json_blocks tidak ada lagi."""
        assert not hasattr(main, '_extract_json_blocks')

    def test_is_header_like_json_removed(self):
        """Fungsi _is_header_like_json tidak ada lagi."""
        assert not hasattr(main, '_is_header_like_json')

    def test_is_response_like_json_removed(self):
        """Fungsi _is_response_like_json tidak ada lagi."""
        assert not hasattr(main, '_is_response_like_json')

    def test_is_header_key_value_line_removed(self):
        """Fungsi _is_header_key_value_line tidak ada lagi."""
        assert not hasattr(main, '_is_header_key_value_line')

    def test_bss_hit_scenarios_removed(self):
        """Konstanta BSS_HIT_SCENARIOS tidak ada lagi."""
        assert not hasattr(main, 'BSS_HIT_SCENARIOS')

    def test_mitra_hit_scenarios_removed(self):
        """Konstanta MITRA_HIT_SCENARIOS tidak ada lagi."""
        assert not hasattr(main, 'MITRA_HIT_SCENARIOS')

    def test_header_keys_removed(self):
        """Konstanta _HEADER_KEYS tidak ada lagi."""
        assert not hasattr(main, '_HEADER_KEYS')

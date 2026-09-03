"""
Unit tests untuk deteksi section header yang FLEKSIBEL (toleran variasi
penulisan antar mitra). Contoh penting: Virtual Account bisa ditulis
'Transfer VA', 'Virtual Account', 'Transfer Virtual Account', dll.
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import detect_section_header


def _section_row(text_b="", text_c=""):
    """Baris section header: kolom E (idx4) & F (idx5) kosong, teks di B/C."""
    r = [None] * 12
    r[1] = text_b
    r[2] = text_c
    return tuple(r)


class TestDeteksiSectionFleksibel:
    def test_virtual_account_berbagai_penulisan(self):
        # Semua varian VA harus terpetakan ke section kanonik "Transfer VA"
        for teks in ["Transfer VA", "Virtual Account",
                     "Transfer Virtual Account", "VA Transfer"]:
            assert detect_section_header(_section_row(teks)) == "Transfer VA", teks

    def test_va_prima(self):
        for teks in ["Transfer VA Prima", "Virtual Account Prima"]:
            assert detect_section_header(_section_row(teks)) == "Transfer VA Prima", teks

    def test_va_bifast(self):
        for teks in ["Transfer VA BI FAST", "Virtual Account BIFAST",
                     "VA Bi-Fast"]:
            assert detect_section_header(_section_row(teks)) == "Transfer VA BI FAST", teks

    def test_section_lain_tetap_benar(self):
        assert detect_section_header(_section_row("Balance Services")) == "Balance Services"
        assert detect_section_header(_section_row("Intrabank Transfer")) == "Intrabank Transfer"
        assert detect_section_header(_section_row("Interbank Transfer")) == "Interbank Transfer"
        assert detect_section_header(_section_row("Interbank Transfer via BI FAST")) == "Interbank Transfer via BI FAST"
        assert detect_section_header(_section_row("RTGS Transfer")) == "RTGS Transfer"
        assert detect_section_header(_section_row("SKNBI Transfer")) == "SKNBI Transfer"

    def test_bukan_section_tidak_terdeteksi(self):
        assert detect_section_header(_section_row("Cashback")) is None
        assert detect_section_header(_section_row("Pengecekan Mutasi")) is None

    def test_baris_data_bukan_section(self):
        # Baris dengan Nomor Kasus (E) terisi bukan section header
        r = [None] * 12
        r[1] = "Transfer VA"
        r[4] = "7.1"   # Nomor Kasus Tes terisi
        assert detect_section_header(tuple(r)) is None

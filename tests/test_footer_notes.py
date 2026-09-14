"""
Unit tests untuk catatan/ketentuan resmi ASPI di akhir dokumen.
Teks ini WAJIB ada dan tidak boleh hilang.
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import build_lampiran_document, FOOTER_NOTES, PRODUCT_PROFILES


def _dummy_lampiran():
    """Buat lampiran_data minimal dengan 1 baris terisi di section pertama."""
    profile = PRODUCT_PROFILES["FT_VA"]
    data = {}
    for name, count in profile["sections"]:
        data[name] = [None] * count
    first = profile["sections"][0][0]
    data[first][0] = {
        "no": 1, "no_full": "1.1", "service": "X", "scenario": "Y",
        "expected_result": "Z", "request": "req", "response": "resp",
        "result": "PASS", "notes": "",
    }
    return data, profile


class TestFooterNotes:
    def test_footer_muncul_di_dokumen(self):
        data, profile = _dummy_lampiran()
        doc = build_lampiran_document(data, profile=profile)
        full_text = "\n".join(p.text for p in doc.paragraphs)
        assert "sekurangnya 1 Pengguna Layanan atas 1 sub API unverified" in full_text
        assert "telah verified, tidak perlu menyampaikan" in full_text

    def test_footer_muncul_sekali_saja(self):
        # Meski ada banyak section, footer hanya sekali di akhir.
        data, profile = _dummy_lampiran()
        # isi baris di section kedua juga
        second = profile["sections"][1][0]
        data[second][0] = {
            "no": 1, "no_full": "2.1", "service": "X", "scenario": "Y",
            "expected_result": "Z", "request": "r", "response": "r",
            "result": "PASS", "notes": "",
        }
        doc = build_lampiran_document(data, profile=profile)
        count = sum(1 for p in doc.paragraphs
                    if "sekurangnya 1 Pengguna Layanan" in p.text)
        assert count == 1

    def test_footer_font_calibri_10(self):
        data, profile = _dummy_lampiran()
        doc = build_lampiran_document(data, profile=profile)
        target = None
        for p in doc.paragraphs:
            if "sekurangnya 1 Pengguna Layanan" in p.text:
                target = p
                break
        assert target is not None
        run = target.runs[0]
        assert run.font.name == "Calibri"
        assert run.font.size.pt == 10

    def test_semua_sub_poin_ada(self):
        # a, b, c, d + 2 butir utama = 6 entri
        assert len(FOOTER_NOTES) == 6
        data, profile = _dummy_lampiran()
        doc = build_lampiran_document(data, profile=profile)
        full_text = "\n".join(p.text for p in doc.paragraphs)
        for penanda in ["a.", "b.", "c.", "d."]:
            assert penanda in full_text

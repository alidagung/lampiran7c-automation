"""Nama Pengguna Layanan & Tanggal Pengujian Lampiran 7C diisi otomatis."""
import io
import os
import sys
from datetime import date
from pathlib import Path

import pytest
from docx import Document

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import main as m

INPUT = Path(__file__).resolve().parent.parent / "input"


def _teks(doc_bytes):
    doc = Document(io.BytesIO(doc_bytes))
    return "\n".join(p.text for p in doc.paragraphs)


def _uat(nama):
    hasil = sorted(INPUT.glob(f"UAT Script*{nama}*.xlsx"))
    if not hasil:
        pytest.skip(f"file contoh {nama} tidak ada")
    return hasil[0]


class TestFormatTanggal7C:
    def test_format(self):
        assert m.format_tanggal_7c(date(2026, 9, 4)) == "04 September 2026"
        assert m.format_tanggal_7c(date(2026, 7, 30)) == "30 Juli 2026"


class TestInfo7C:
    def test_diisi_dari_parameter(self):
        profile = m.PRODUCT_PROFILES["FT_VA"]
        data = {name: [None] * count for name, count in profile["sections"]}
        data[profile["sections"][0][0]][0] = {
            "no": 1, "no_full": "1.1", "service": "X", "scenario": "Y",
            "expected_result": "Z", "request": "req", "response": "resp",
            "result": "PASS", "notes": "",
        }
        doc = m.build_lampiran_document(data, profile=profile, pengguna_layanan="PT Contoh Mitra",
                                        tanggal_pengujian=date(2026, 7, 30))
        teks = "\n".join(p.text for p in doc.paragraphs)
        assert "Nama Pengguna Layanan : PT Contoh Mitra" in teks
        assert "Tanggal Pengujian     : 30 Juli 2026" in teks

    @pytest.mark.parametrize("nama, pt", [
        ("Pipo", "PT Fintech Bina Bangsa"),
        ("Kirimo", "PT Sender Integrasi Digital"),
    ])
    def test_otomatis_dari_judul(self, nama, pt):
        src = _uat(nama)
        out = m.convert_uat_to_lampiran_bytes(src.read_bytes(), tanggal_pengujian=date(2026, 9, 4))
        teks = _teks(out[0])
        assert f"Nama Pengguna Layanan : {pt}" in teks
        assert "Tanggal Pengujian     : 04 September 2026" in teks

    def test_tanpa_tanggal_pakai_hari_ini(self):
        src = _uat("Pipo")
        teks = _teks(m.convert_uat_to_lampiran_bytes(src.read_bytes())[0])
        assert f"Tanggal Pengujian     : {m.format_tanggal_7c(date.today())}" in teks

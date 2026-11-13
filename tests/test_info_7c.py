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


class TestSiapDitesButuhKonfirmasi:
    """Siap dites / Butuh konfirmasi: log -> Request/Response; catatan -> Notes."""

    LOG = ('URL:\nhttps://x/api\n\nHeaders:\nA: b\n\nRequest Body:\n{"a": 1}\n\n'
           'Response:\n{"ok": 1}')

    def _map(self, hasil, remarks):
        uat = {"Balance Services": [{
            "nomor_kasus_tes": "1.1", "langkah_tes": "x", "hasil_aktual": hasil,
            "hasil_diharapkan": "", "nama_modul": "Balance Inquiry", "remarks": remarks}]}
        data = m.map_uat_to_lampiran(uat, profile=m.PRODUCT_PROFILES["FT_VA"])
        return next(r for rows in data.values() for r in rows if r)

    @pytest.mark.parametrize("hasil", ["Siap dites", "Butuh konfirmasi"])
    def test_ada_log_dilampirkan_ke_request_response(self, hasil):
        r = self._map(hasil, self.LOG)
        assert "https://x/api" in r["request"] and '"ok"' in r["response"]
        assert r["notes"] == "" and r["result"] == "N/A"

    @pytest.mark.parametrize("hasil", ["Siap dites", "Butuh konfirmasi"])
    def test_hanya_catatan_masuk_notes(self, hasil):
        r = self._map(hasil, "Menunggu update fixing dari vendor.")
        assert r["request"] == "" and r["response"] == ""
        assert r["notes"] == "Menunggu update fixing dari vendor."
        assert r["result"] == "N/A"

    def test_remark_kosong_notes_kosong(self):
        r = self._map("Siap dites", "")
        assert (r["request"], r["response"], r["notes"]) == ("", "", "")


class TestNotesTidakDites:
    def _map(self, hasil, remarks):
        return TestSiapDitesButuhKonfirmasi._map(None, hasil, remarks)

    @pytest.mark.parametrize("hasil", ["Tidak dites", "Belum dites", "Siap dites", "Butuh konfirmasi"])
    def test_remark_kosong_notes_ikut_kosong(self, hasil):
        r = self._map(hasil, "")
        assert r["notes"] == "" and r["result"] == "N/A"

    def test_remark_ada_dipakai_apa_adanya(self):
        assert self._map("Tidak dites", "Tidak dites karena X.")["notes"] == "Tidak dites karena X."

    def test_notifikasi_remark_kosong(self):
        w = m.detect_anomalies({"nomor_kasus_tes": "7.1", "hasil_aktual": "Tidak dites", "remarks": ""})
        assert len(w) == 1 and "7.1" in w[0] and "Remarks kosong" in w[0]
        assert m.detect_anomalies({"nomor_kasus_tes": "7.1", "hasil_aktual": "Tidak dites",
                                   "remarks": "alasan"}) == []


class TestAlasanLevelApi:
    """Remarks case x.1 (API tidak dites) jadi Notes case lain yang Remarks-nya kosong."""

    def _rows(self, first_hasil, first_remark):
        return [
            {"nomor_kasus_tes": "1.1", "langkah_tes": "a", "hasil_aktual": first_hasil,
             "hasil_diharapkan": "", "nama_modul": "Any Service", "remarks": first_remark},
            {"nomor_kasus_tes": "1.2", "langkah_tes": "b", "hasil_aktual": "Belum dites",
             "hasil_diharapkan": "", "nama_modul": "Any Service", "remarks": ""},
            {"nomor_kasus_tes": "1.3", "langkah_tes": "c", "hasil_aktual": "Tidak dites",
             "hasil_diharapkan": "", "nama_modul": "Balance Inquiry", "remarks": "Alasan sendiri"},
        ]

    def _notes(self, rows):
        data, w = m.map_uat_to_lampiran({"Balance Services": rows}, collect_warnings=True,
                                        profile=m.PRODUCT_PROFILES["FT_VA"])
        r = data["API Balance Inquiry"]
        return [x["notes"] for x in r[:3]], w

    def test_alasan_x1_dipakai_untuk_remark_kosong(self):
        alasan = "Tidak dites karena Mitra tidak menggunakan Balance Services"
        notes, w = self._notes(self._rows("Belum dites", alasan))
        assert notes == [alasan, alasan, "Alasan sendiri"]
        assert not [x for x in w if "Remarks kosong" in x]

    def test_x1_berhasil_tidak_diwariskan(self):
        notes, _ = self._notes(self._rows("Berhasil", 'URL:\nhttps://x\n\nResponse:\n{"a":1}'))
        assert notes[1] == "" and notes[2] == "Alasan sendiri"

    def test_x1_tanpa_alasan_notes_kosong_dan_notifikasi(self):
        notes, w = self._notes(self._rows("Tidak dites", ""))
        assert notes[:2] == ["", ""]
        assert len([x for x in w if "Remarks kosong" in x]) == 2

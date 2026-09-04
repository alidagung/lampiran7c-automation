"""
Unit tests untuk deteksi anomali (detect_anomalies).
Memastikan sistem peringatan mendeteksi data abnormal TANPA mengubah data.
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import detect_anomalies


def _row(kasus, hasil, remarks):
    return {
        "nomor_kasus_tes": kasus,
        "hasil_aktual": hasil,
        "remarks": remarks,
    }


REMARKS_LENGKAP = (
    "URL:\nhttps://api.test/x\n\n"
    "Headers:\nAuthorization: Bearer x\n\n"
    "Request Body:\n{\"a\":1}\n\n"
    "Response:\n{\"ok\":true}"
)


class TestDetectAnomalies:
    def test_data_lengkap_tidak_ada_warning(self):
        assert detect_anomalies(_row("1.1", "Berhasil", REMARKS_LENGKAP)) == []

    def test_remarks_kosong_saat_berhasil(self):
        w = detect_anomalies(_row("1.2", "Berhasil", ""))
        assert len(w) == 1
        assert "Remarks kosong" in w[0]

    def test_response_json_tidak_valid(self):
        remarks = (
            "URL:\nhttps://x\n\nHeaders:\nA: b\n\n"
            "Request Body:\n{\"a\":1}\n\nResponse:\n{ini bukan json}"
        )
        w = detect_anomalies(_row("2.5", "Berhasil", remarks))
        assert any("Response Body bukan JSON valid" in x for x in w)

    def test_request_body_json_tidak_valid(self):
        remarks = (
            "URL:\nhttps://x\n\nHeaders:\nA: b\n\n"
            "Request Body:\n{rusak}\n\nResponse:\n{\"ok\":true}"
        )
        w = detect_anomalies(_row("2.6", "Berhasil", remarks))
        assert any("Request Body bukan JSON valid" in x for x in w)

    def test_bagian_response_hilang(self):
        remarks = "URL:\nhttps://x\n\nHeaders:\nA: b\n\nRequest Body:\n{\"a\":1}"
        w = detect_anomalies(_row("3.2", "Berhasil", remarks))
        assert any("Response tidak ditemukan" in x for x in w)

    def test_tidak_dites_tidak_ada_warning(self):
        # Baris yang tidak dites tidak seharusnya memicu peringatan
        assert detect_anomalies(_row("4.1", "Tidak dites", "")) == []

    def test_belum_dites_tidak_ada_warning(self):
        assert detect_anomalies(_row("4.2", "Belum dites", "")) == []

    def test_warning_menyebut_nomor_kasus(self):
        w = detect_anomalies(_row("7.15", "Berhasil", ""))
        assert "7.15" in w[0]



class TestDeteksiTambahan:
    """Deteksi yang lebih teliti (tanpa mengubah data)."""

    def test_request_body_kurung_tidak_lengkap(self):
        # Kurung penutup '}' hilang -> JSON tidak valid
        remarks = (
            "URL:\nhttps://x\n\nHeaders:\nA: b\n\n"
            'Request Body:\n{\n"accountNo": "123",\n  "ref": "abc"\n\n'
            'Response:\n{"ok":true}'
        )
        w = detect_anomalies(_row("1.1", "Berhasil", remarks))
        assert any("Request Body bukan JSON valid" in x for x in w)

    def test_response_kosong_saat_berhasil(self):
        remarks = (
            "URL:\nhttps://x\n\nHeaders:\nA: b\n\n"
            'Request Body:\n{"a":1}\n\nResponse:\n'
        )
        w = detect_anomalies(_row("2.2", "Berhasil", remarks))
        assert any("Response Body kosong" in x for x in w)

    def test_header_tidak_wajar(self):
        # Ada item header yang bukan Key=Value / Key: Value
        remarks = (
            "URL:\nhttps://x\n\nHeaders:\n[\n  \"barisrusaktanpaseparator\"\n]\n\n"
            'Request Body:\n{"a":1}\n\nResponse:\n{"ok":true}'
        )
        w = detect_anomalies(_row("3.3", "Berhasil", remarks))
        assert any("Header tidak wajar" in x for x in w)

    def test_data_normal_tidak_ada_peringatan_baru(self):
        # accountNo kosong TAPI JSON tetap valid -> TIDAK dianggap masalah
        remarks = (
            "URL:\nhttps://x\n\nHeaders:\nContent-Type: application/json\n\n"
            'Request Body:\n{"accountNo": "", "ref": "abc"}\n\n'
            'Response:\n{"responseCode":"200"}'
        )
        assert detect_anomalies(_row("4.4", "Berhasil", remarks)) == []



class TestDisplayNoPenomoranLampiran:
    """
    Peringatan harus memakai nomor kasus SESUAI penomoran Lampiran 7C
    (parameter display_no), bukan nomor asli file UAT. Penting agar laporan
    ke ASPI konsisten dengan kolom No pada dokumen.
    """

    def test_display_no_dipakai_di_pesan(self):
        # Nomor asli 7.11, tapi di Lampiran 7C tampil sebagai 6.11
        w = detect_anomalies(_row("7.11", "Berhasil", ""), display_no="6.11")
        assert len(w) == 1
        assert "6.11" in w[0]
        assert "7.11" not in w[0]

    def test_tanpa_display_no_pakai_nomor_asli(self):
        # Backward compatible: tanpa display_no -> pakai nomor_kasus_tes
        w = detect_anomalies(_row("7.11", "Berhasil", ""))
        assert "7.11" in w[0]

    def test_display_no_pada_response_invalid(self):
        remarks = (
            "URL:\nhttps://x\n\nHeaders:\nA: b\n\n"
            "Request Body:\n{\"a\":1}\n\nResponse:\n{ini bukan json}"
        )
        w = detect_anomalies(_row("7.2", "Berhasil", remarks), display_no="6.2")
        assert any("6.2" in x for x in w)
        assert all("7.2" not in x for x in w)

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

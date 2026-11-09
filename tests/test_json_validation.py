"""
Aturan: isi JSON tidak boleh diubah; spasi/enter tidak divalidasi; JSON rusak
tetap tampil persis seperti sumber, sedangkan notifikasi tetap diberikan.
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import (
    _compress_json, _pretty_json, _compress_response, _is_valid_json,
    detect_anomalies,
)


class TestIsiJsonTidakBerubah:
    def test_angka_dan_spasi_dalam_string_dipertahankan(self):
        t = '{"a": 10.50, "b":  "x   y", "e": 1E5, "c": [], "d": {}}'
        assert _compress_json(t) == '{"a":10.50,"b":"x   y","e":1E5,"c":[],"d":{}}'
        assert "10.50" in _pretty_json(t) and "1E5" in _pretty_json(t)

    def test_json_rusak_dikembalikan_persis_seperti_sumber(self):
        rusak = '{\n  "a": 1\n  "b": 2\n}'   # kurang koma
        assert _compress_json(rusak) == rusak
        assert _pretty_json(rusak) == rusak

    def test_enter_dalam_string_tidak_dianggap_error(self):
        assert _is_valid_json('{"a": "baris1\nbaris2"}')

    def test_response_raw_http_tetap_satu_baris(self):
        raw = "HTTP/1.1 401 \nContent-Length: 66\nContent-Type: application/json"
        assert _compress_response(raw) == "HTTP/1.1 401 Content-Length: 66 Content-Type: application/json"


def _row(remarks):
    return {"nomor_kasus_tes": "9.9", "hasil_aktual": "Berhasil", "remarks": remarks}


class TestNotifikasiJson:
    def test_pesan_menyebut_baris_dan_kolom(self):
        remarks = ('URL:\nhttps://x\n\nHeaders:\nA=b\n\nRequest Body:\n{"a":1}\n\n'
                   'Response:\n{\n  "x": 1\n  "y": 2\n}')
        w = detect_anomalies(_row(remarks))
        assert any("Response Body bukan JSON valid" in x and "baris" in x and "kolom" in x for x in w)

    def test_json_menempel_setelah_request_body_dilaporkan(self):
        remarks = ('URL:\nhttps://x\n\nHeaders:\nA=b\n\nRequest Body:\n{"a":1}\n\n'
                   'Response Body{\n  "responseCode": "400"\n}')
        w = detect_anomalies(_row(remarks))
        assert any("teks/blok tambahan setelah JSON pertama" in x for x in w)
        assert any("Response Body{" in x for x in w)

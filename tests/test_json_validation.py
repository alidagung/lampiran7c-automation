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

    def test_json_rusak_dirapatkan_tanpa_ubah_isi(self):
        rusak = '{\n  "a": 1\n  "b": 2\n}'   # kurang koma
        assert _compress_json(rusak) == '{"a":1 "b":2}'      # enter hilang, isi tetap
        assert _pretty_json(rusak) == rusak

    def test_json_rusak_isi_string_tidak_disentuh(self):
        rusak = '{ "a": "x   y",\n "b" : 1 } {"c":2}'
        assert _compress_json(rusak) == '{"a":"x   y","b":1}{"c":2}'

    def test_bukan_json_dibiarkan(self):
        assert _compress_json("Tidak ada\nresponse") == "Tidak ada\nresponse"

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

    def test_label_response_tanpa_titik_dua_dikenali(self):
        # "Response Body{" (tanpa ':') harus dikenali sbg label -> tidak ada
        # peringatan JSON menempel, dan Response terpisah dari Request Body.
        remarks = ('URL:\nhttps://x\n\nHeaders:\nA=b\n\nRequest Body:\n{"a":1}\n\n'
                   'Response Body{\n  "responseCode": "400"\n}')
        w = detect_anomalies(_row(remarks))
        assert not any("bukan JSON valid" in x for x in w)
        assert not any("Response Body kosong" in x or "Response tidak ditemukan" in x for x in w)

    def test_dua_blok_json_tanpa_label_tetap_dilaporkan(self):
        remarks = ('URL:\nhttps://x\n\nHeaders:\nA=b\n\nRequest Body:\n{"a":1}\n\n'
                   '{"responseCode": "400"}')
        w = detect_anomalies(_row(remarks))
        assert any("teks/blok tambahan setelah JSON pertama" in x for x in w)


class TestLabelResponseTanpaTitikDua:
    def test_response_masuk_kolom_response(self):
        from main import split_request_response, split_request_response_qris
        va = ('URL:\nhttps://x\n\nHeaders:\nA=b\n\nRequest Body:\n{"a":1}\n\n'
              'Response Body{\n  "responseCode": "4002402"\n}')
        req, resp = split_request_response(va)
        assert "4002402" in resp and "4002402" not in req
        qr = ('URL: https://x/notify\nContent-Type: application/json\n\n{"a":1}\n\n'
              'Response {\n "responseCode": "2005200"\n}')
        req2, resp2 = split_request_response_qris(qr)
        assert "2005200" in resp2 and "2005200" not in req2

    def test_value_json_tidak_berubah(self):
        from main import split_request_response
        va = ('URL:\nhttps://x\n\nHeaders:\nA=b\n\nRequest Body:\n{"a":1}\n\n'
              'Response Body\n{"responseMessage": "Response Body{ tetap"}')
        _, resp = split_request_response(va)
        assert 'Response Body{ tetap' in resp

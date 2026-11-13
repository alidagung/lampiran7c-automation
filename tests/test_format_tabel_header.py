"""
Format Remark mitra (mis. PT Solusi Pembayaran Elektronik): label tanpa titik
dua, blok 'Request Method:', dan header berbentuk tabel 'Name  Value'.
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import split_request_response, detect_anomalies, _compress_headers_inline

REMARKS = (
    "Request URL:\nhttps://ob-sandbox.banksampoerna.co.id/snap/v1.2/balance-inquiry\n\n"
    "Request Method:\nPOST\n\n"
    "Request Headers\n"
    "Name        Value\n"
    "content-type        application/json\n"
    "X-PARTNER-ID        x8ir2mpkj477amer19biu4kn2l8cwlcw\n"
    "Authorization        Bearer test\n"
    "Cookie        Path=/\n\n"
    "Request Body\n{\n  \"accountNo\": \"1023653296\"\n}\n\n"
    "Response Body\n\n{\n  \"responseCode\": \"4011101\"\n}"
)


class TestFormatTabelHeader:
    def test_header_tabel_jadi_key_value_sesuai_contoh(self):
        req, _ = split_request_response(REMARKS)
        assert ('["content-type=application/json",'
                '"X-PARTNER-ID=x8ir2mpkj477amer19biu4kn2l8cwlcw",'
                '"Authorization=Bearer test",'
                '"Cookie=Path=/"]') in req
        assert "Name" not in req.split("Header Request:")[1].split("Request Body:")[0]

    def test_url_body_response_di_tempat_yang_benar(self):
        req, resp = split_request_response(REMARKS)
        assert "URL Endpoint:\nhttps://ob-sandbox.banksampoerna.co.id/snap/v1.2/balance-inquiry" in req
        body = req.split("Request Body:")[1]
        assert '"accountNo": "1023653296"' in body and "POST" not in body   # bersusun ke bawah
        assert resp == 'Response Body:\n{"responseCode":"4011101"}'

    def test_tidak_ada_peringatan_palsu(self):
        w = detect_anomalies({"nomor_kasus_tes": "1.1", "hasil_aktual": "Berhasil", "remarks": REMARKS})
        assert w == []

    def test_uat_result_header_inline(self):
        out = _compress_headers_inline("Name        Value\ncontent-type        application/json\nCookie        Path=/")
        assert out == "[content-type=application/json, Cookie=Path=/]"

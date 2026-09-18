"""
Unit tests untuk parser QRIS (format raw HTTP) -> URL Endpoint, Header
Request, dan Request Body terdeteksi otomatis dan disusun seperti VA.
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import (
    split_request_response_qris,
    _parse_qris_request_block,
    detect_anomalies,
)


REMARKS_QRMPM = (
    "Request: \n"
    "POST /snap-qris/v1.1/qr/qr-mpm-generate HTTP/1.1\n"
    "Host: ob-sandbox.banksampoerna.co.id\n"
    "Authorization: Bearer xxx\n"
    "CHANNEL-ID: CH001\n"
    "Content-Type: application/json\n"
    "X-PARTNER-ID: KIRIMO\n"
    "\n"
    '{"partnerReferenceNo":"123","merchantId":"9360"}\n'
    "\n"
    "Response:\n"
    "HTTP/1.1 200 \n"
    "Content-Type: application/json\n"
    "\n"
    '{"responseCode":"2004700"}'
)


class TestParseQrisRequestBlock:
    def test_url_persis_request_line_dan_host(self):
        # URL Endpoint ditampilkan APA ADANYA: request line + baris Host.
        req = REMARKS_QRMPM.split("Response:")[0]
        url, headers, body = _parse_qris_request_block(req)
        assert url == (
            "POST /snap-qris/v1.1/qr/qr-mpm-generate HTTP/1.1\n"
            "Host: ob-sandbox.banksampoerna.co.id"
        )

    def test_host_tidak_masuk_header(self):
        # Host hanya di URL Endpoint, TIDAK di Header Request.
        req = REMARKS_QRMPM.split("Response:")[0]
        _, headers, _ = _parse_qris_request_block(req)
        joined = "\n".join(headers)
        assert "Authorization: Bearer xxx" in joined
        assert "CHANNEL-ID: CH001" in joined
        assert "X-PARTNER-ID: KIRIMO" in joined
        assert not any(h.lower().startswith("host") for h in headers)

    def test_body_terdeteksi(self):
        req = REMARKS_QRMPM.split("Response:")[0]
        _, _, body = _parse_qris_request_block(req)
        assert '"partnerReferenceNo"' in body

    def test_url_label_eksplisit(self):
        req = "URL: https://api-staging.kirimo.dev/provider/v1.0/qr/qr-mpm-notify\nContent-Type: application/json\n\n{\"a\":1}"
        url, headers, body = _parse_qris_request_block(req)
        assert url == "https://api-staging.kirimo.dev/provider/v1.0/qr/qr-mpm-notify"
        assert body == '{"a":1}'


class TestSplitQris:
    def test_request_terformat_seperti_va(self):
        req, resp = split_request_response_qris(REMARKS_QRMPM)
        assert "URL Endpoint:" in req
        assert "Header Request:" in req
        assert "Request Body:" in req
        # Header disusun sebagai array Key=Value seperti VA
        assert "Authorization=Bearer xxx" in req

    def test_response_dipisah_dan_compress(self):
        req, resp = split_request_response_qris(REMARKS_QRMPM)
        assert resp.startswith("Response Body:")
        assert "Response Body:" not in req


REMARKS_QRIS_BERLABEL = (
    "Url : https://ob-sandbox.banksampoerna.co.id/snap-qris/v1.1/qr/qr-mpm-generate\n\n"
    "Request headers: {\n"
    '  "Authorization": ["Bearer DUMMY_TOKEN"],\n'
    '  "CHANNEL-ID": ["CH001"],\n'
    '  "Content-Type": ["application/json"],\n'
    '  "X-PARTNER-ID": ["XALLURE"]\n'
    "}\n"
    "Request body: {\n"
    '  "merchantId": "9360052300000003681",\n'
    '  "partnerReferenceNo": "148IYTKXC2WATM822UVFT084B"\n'
    "}\n"
    "Response body: {\n"
    '  "responseCode": "4014701",\n'
    '  "responseMessage": "Invalid token (B2B)"\n'
    "}"
)


class TestQrisFormatBerlabel:
    """
    Mitra QRIS lain (mis. XALLURE) memakai format BERLABEL dengan header JSON
    object. Header HARUS masuk ke Header Request, BUKAN ke Request Body.
    """

    def test_header_json_masuk_header_bukan_body(self):
        req, resp = split_request_response_qris(REMARKS_QRIS_BERLABEL)
        # Header Request berisi item header
        assert "Header Request:" in req
        assert "Authorization=Bearer DUMMY_TOKEN" in req
        assert "X-PARTNER-ID=XALLURE" in req
        # Request Body berisi field body, BUKAN Authorization
        assert "Request Body:" in req
        # Pastikan Authorization TIDAK berada di blok Request Body
        body_part = req.split("Request Body:")[1]
        assert "Authorization" not in body_part
        assert "merchantId" in body_part

    def test_url_terbaca(self):
        req, resp = split_request_response_qris(REMARKS_QRIS_BERLABEL)
        assert "qr-mpm-generate" in req
        assert "URL Endpoint:" in req

    def test_response_dipisah(self):
        req, resp = split_request_response_qris(REMARKS_QRIS_BERLABEL)
        assert resp.startswith("Response Body:")
        assert "4014701" in resp


class TestDetectAnomaliesQris:
    def test_qris_data_lengkap_tidak_ada_warning_json(self):
        # Request Body JSON valid; response raw HTTP (tidak dicek JSON di QRIS)
        w = detect_anomalies(
            {"nomor_kasus_tes": "3.6", "hasil_aktual": "Berhasil", "remarks": REMARKS_QRMPM},
            request_mode="qris",
        )
        # Tidak ada warning "bukan JSON valid" untuk Response di mode QRIS
        assert not any("Response Body bukan JSON valid" in x for x in w)
        # URL, Header, Request Body, Response semuanya terdeteksi -> tak ada "tidak ditemukan"
        assert not any("tidak ditemukan" in x for x in w)

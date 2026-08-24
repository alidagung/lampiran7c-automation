"""
Unit tests untuk fungsi parsing kolom Remarks di main.py.
Menguji kedua format: 'Mitra HIT' dan 'BSS YANG HIT'.
Termasuk test untuk content-detection approach (tanpa keyword/label).
"""

import sys
import os

# Tambahkan root project ke path agar bisa import main
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import (
    parse_remarks_mitra_hit,
    parse_remarks_bss_hit,
    parse_remarks,
    _extract_url_from_text,
    _extract_json_blocks,
    _is_header_like_json,
    _is_response_like_json,
    _parse_by_content_detection,
    _parse_with_labels,
)


class TestParseMitraHit:
    """Test parsing format 'Mitra HIT' (Skenario 2, 3, 4, 5, 6)."""

    def test_standard_format(self):
        """Test format standar dengan semua label."""
        remarks = """URL:
https://api.example.com/v1/transfer

Header:
{"Content-Type": "application/json", "Authorization": "Bearer token123"}

RequestBody:
{"amount": 50000, "destination": "1234567890"}

Response:
{"status": "success", "transactionId": "TXN001"}"""

        result = parse_remarks_mitra_hit(remarks)

        assert result is not None
        assert "https://api.example.com/v1/transfer" in result['url']
        assert "Content-Type" in result['headers']
        assert "amount" in result['request_body']
        assert "success" in result['response']

    def test_url_without_label(self):
        """Test ketika Remarks dimulai langsung dengan URL tanpa label."""
        remarks = "https://api.example.com/v1/balance"

        result = parse_remarks_mitra_hit(remarks)

        assert result is not None
        assert "https://api.example.com/v1/balance" in result['url']

    def test_url_http_without_label(self):
        """Test ketika Remarks dimulai dengan http:// URL tanpa label."""
        remarks = """http://api.example.com/v1/balance
{"status": "ok"}"""

        result = parse_remarks_mitra_hit(remarks)

        assert result is not None
        assert "http://api.example.com/v1/balance" in result['url']

    def test_empty_remarks(self):
        """Test dengan Remarks kosong."""
        assert parse_remarks_mitra_hit("") is None
        assert parse_remarks_mitra_hit(None) is None
        assert parse_remarks_mitra_hit("   ") is None

    def test_partial_data(self):
        """Test dengan hanya beberapa field yang ada."""
        remarks = """URL:
https://api.example.com/v1/rtgs

Response:
{"status": "failed", "error": "insufficient_funds"}"""

        result = parse_remarks_mitra_hit(remarks)

        assert result is not None
        assert "https://api.example.com/v1/rtgs" in result['url']
        assert "insufficient_funds" in result['response']

    def test_multiline_request_body(self):
        """Test dengan request body JSON yang multi-line."""
        remarks = """URL:
https://api.example.com/v1/transfer

Header:
{"Authorization": "Bearer xyz"}

RequestBody:
{
    "amount": 100000,
    "sourceAccount": "111222333",
    "destAccount": "444555666",
    "description": "Transfer test"
}

Response:
{"status": "success"}"""

        result = parse_remarks_mitra_hit(remarks)

        assert result is not None
        assert "amount" in result['request_body']
        assert "sourceAccount" in result['request_body']
        assert "destAccount" in result['request_body']
        assert "success" in result['response']

    def test_headers_variation(self):
        """Test dengan variasi penulisan 'Headers' (dengan s)."""
        remarks = """URL:
https://api.example.com/v1/sknbi

Headers:
{"Content-Type": "application/json"}

RequestBody:
{"data": "test"}

Response:
{"ok": true}"""

        result = parse_remarks_mitra_hit(remarks)

        assert result is not None
        assert "Content-Type" in result['headers']


class TestParseBssHit:
    """Test parsing format 'BSS YANG HIT' (Skenario 7, 8, 9)."""

    def test_standard_format(self):
        """Test format standar BSS HIT."""
        remarks = """- url: https://api.example.com/v1/va/create
- headers: {"Content-Type": "application/json", "X-API-Key": "key123"}
- Request Body: {"vaNumber": "8800123456", "amount": 75000}
- Response: {"status": "00", "vaNumber": "8800123456"}"""

        result = parse_remarks_bss_hit(remarks)

        assert result is not None
        assert "https://api.example.com/v1/va/create" in result['url']
        assert "X-API-Key" in result['headers']
        assert "vaNumber" in result['request_body']
        assert '"status": "00"' in result['response']

    def test_lowercase_body_and_response(self):
        """Test dengan variasi 'body' dan 'response' lowercase."""
        remarks = """- url: https://api.example.com/v1/va/inquiry
- headers: {"Authorization": "Basic abc123"}
- body: {"accountNumber": "9900111222"}
- response: {"status": "00", "balance": 500000}"""

        result = parse_remarks_bss_hit(remarks)

        assert result is not None
        assert "https://api.example.com/v1/va/inquiry" in result['url']
        assert "Authorization" in result['headers']
        assert "accountNumber" in result['request_body']
        assert "balance" in result['response']

    def test_empty_remarks(self):
        """Test dengan Remarks kosong."""
        assert parse_remarks_bss_hit("") is None
        assert parse_remarks_bss_hit(None) is None
        assert parse_remarks_bss_hit("   ") is None

    def test_partial_data(self):
        """Test dengan hanya beberapa field yang ada."""
        remarks = """- url: https://api.example.com/v1/va/status
- Response: {"status": "00", "message": "Active"}"""

        result = parse_remarks_bss_hit(remarks)

        assert result is not None
        assert "https://api.example.com/v1/va/status" in result['url']
        assert "Active" in result['response']

    def test_multiline_json_values(self):
        """Test dengan JSON values yang panjang."""
        remarks = """- url: https://api.example.com/v1/va/payment
- headers: {"Content-Type": "application/json", "X-Request-ID": "req-123-456-789", "Authorization": "Bearer longtoken"}
- Request Body: {"virtualAccountNumber": "8800999888777", "amount": 250000, "customerName": "John Doe", "description": "Payment for order #12345"}
- Response: {"responseCode": "00", "responseMessage": "Success", "transactionId": "TXN-2024-001"}"""

        result = parse_remarks_bss_hit(remarks)

        assert result is not None
        assert "virtualAccountNumber" in result['request_body']
        assert "responseCode" in result['response']

    def test_no_matching_pattern(self):
        """Test dengan text yang tidak sesuai format BSS HIT."""
        remarks = "Ini adalah catatan biasa tanpa format khusus"

        result = parse_remarks_bss_hit(remarks)

        assert result is None


class TestParseRemarks:
    """Test fungsi parse_remarks utama yang memilih parser berdasarkan skenario."""

    def test_mitra_hit_scenario_2(self):
        """Test bahwa skenario 2 menggunakan parser Mitra HIT."""
        remarks = """URL:
https://api.example.com/v1/intrabank

Header:
{"Auth": "token"}

RequestBody:
{"amount": 10000}

Response:
{"status": "success"}"""

        result = parse_remarks(remarks, "2")

        assert result is not None
        assert "intrabank" in result['url']

    def test_mitra_hit_scenario_3(self):
        """Test bahwa skenario 3 menggunakan parser Mitra HIT."""
        remarks = """URL:
https://api.example.com/v1/interbank

Response:
{"ok": true}"""

        result = parse_remarks(remarks, "3")

        assert result is not None
        assert "interbank" in result['url']

    def test_bss_hit_scenario_7(self):
        """Test bahwa skenario 7 menggunakan parser BSS HIT."""
        remarks = """- url: https://api.example.com/v1/va
- headers: {"key": "value"}
- body: {"data": "test"}
- response: {"status": "00"}"""

        result = parse_remarks(remarks, "7")

        assert result is not None
        assert "va" in result['url']

    def test_bss_hit_scenario_8(self):
        """Test bahwa skenario 8 menggunakan parser BSS HIT."""
        remarks = """- url: https://api.example.com/v1/va-prima
- headers: {}
- Request Body: {}
- Response: {"status": "00"}"""

        result = parse_remarks(remarks, "8")

        assert result is not None
        assert "va-prima" in result['url']

    def test_bss_hit_scenario_9(self):
        """Test bahwa skenario 9 menggunakan parser BSS HIT."""
        remarks = """- url: https://api.example.com/v1/va-bifast
- headers: {"token": "abc"}
- body: {"amount": 100}
- response: {"done": true}"""

        result = parse_remarks(remarks, "9")

        assert result is not None
        assert "va-bifast" in result['url']

    def test_scenario_1_fallback(self):
        """Test bahwa skenario 1 (Balance) mencoba kedua parser."""
        remarks = """URL:
https://api.example.com/v1/balance

Response:
{"balance": 1000000}"""

        result = parse_remarks(remarks, "1")

        assert result is not None
        assert "balance" in result['url']

    def test_empty_remarks_all_scenarios(self):
        """Test bahwa semua skenario return None untuk Remarks kosong."""
        for scenario in ["1", "2", "3", "4", "5", "6", "7", "8", "9"]:
            assert parse_remarks("", scenario) is None
            assert parse_remarks(None, scenario) is None


class TestRealWorldFormatA:
    """
    Test parsing Format A (Mitra HIT - Skenario 2,3,4,5,6) dengan data
    yang sesuai format asli dari UAT Script Excel.
    Headers dalam format Key: Value (bukan JSON).
    """

    def test_real_format_a_full(self):
        """Test format asli Mitra HIT dengan headers key:value."""
        remarks = """URL:
https://ob-sandbox.banksampoerna.co.id/snap/v1.2/account-inquiry-internal

Header:
X-TIMESTAMP: 2025-07-30T14:28:56+07:00
X-PARTNER-ID: 1lpm2oeabzo69qfobw0i096s9o4scin5
X-SIGNATURE: TIrc3mz+abc123
Authorization: Bearer 67248JOH
X-EXTERNAL-ID: tXid1753860536066
CHANNEL-ID: 11101
Content-Type: application/json; charset=utf-8

RequestBody:
{"partnerReferenceNo":"900009","beneficiaryAccountNo":"0111901001"}

Response:
{"responseCode":"4011501","responseMessage":"Invalid token (B2B)"}"""

        result = parse_remarks_mitra_hit(remarks)

        assert result is not None
        assert "https://ob-sandbox.banksampoerna.co.id/snap/v1.2/account-inquiry-internal" in result['url']
        assert "X-TIMESTAMP" in result['headers']
        assert "X-PARTNER-ID" in result['headers']
        assert "X-SIGNATURE" in result['headers']
        assert "Authorization" in result['headers']
        assert "CHANNEL-ID" in result['headers']
        assert "Content-Type" in result['headers']
        assert "partnerReferenceNo" in result['request_body']
        assert "responseCode" in result['response']
        assert "4011501" in result['response']

    def test_real_format_a_via_parse_remarks(self):
        """Test that parse_remarks routes correctly for scenario 2."""
        remarks = """URL:
https://ob-sandbox.banksampoerna.co.id/snap/v1.2/account-inquiry-internal

Header:
X-TIMESTAMP: 2025-07-30T14:28:56+07:00
X-PARTNER-ID: 1lpm2oeabzo69qfobw0i096s9o4scin5
Authorization: Bearer 67248JOH
Content-Type: application/json; charset=utf-8

RequestBody:
{"partnerReferenceNo":"900009"}

Response:
{"responseCode":"4011501","responseMessage":"Invalid token (B2B)"}"""

        result = parse_remarks(remarks, "2")

        assert result is not None
        assert "ob-sandbox.banksampoerna.co.id" in result['url']
        assert "X-TIMESTAMP" in result['headers']
        assert "partnerReferenceNo" in result['request_body']
        assert "responseCode" in result['response']

    def test_format_a_no_url_label(self):
        """Test Format A tanpa label URL: (langsung URL)."""
        remarks = """https://ob-sandbox.banksampoerna.co.id/snap/v1.2/transfer

Header:
X-TIMESTAMP: 2025-07-30T14:28:56+07:00
Authorization: Bearer token123
Content-Type: application/json

RequestBody:
{"amount": 50000}

Response:
{"responseCode":"2001700","responseMessage":"Successful"}"""

        result = parse_remarks_mitra_hit(remarks)

        assert result is not None
        assert "ob-sandbox.banksampoerna.co.id" in result['url']
        assert "X-TIMESTAMP" in result['headers']
        assert "amount" in result['request_body']
        assert "Successful" in result['response']

    def test_format_a_no_labels_at_all(self):
        """Test content detection when there are NO labels at all."""
        remarks = """https://ob-sandbox.banksampoerna.co.id/snap/v1.2/transfer
X-TIMESTAMP: 2025-07-30T14:28:56+07:00
X-PARTNER-ID: abc123
Authorization: Bearer token123
Content-Type: application/json
{"partnerReferenceNo":"900009","amount":50000}
{"responseCode":"2001700","responseMessage":"Successful"}"""

        result = parse_remarks_mitra_hit(remarks)

        assert result is not None
        assert "ob-sandbox.banksampoerna.co.id" in result['url']


class TestRealWorldFormatB:
    """
    Test parsing Format B (BSS HIT - Skenario 7,8,9) dengan data
    yang sesuai format asli dari UAT Script Excel.
    """

    def test_real_format_b_full(self):
        """Test format asli BSS HIT."""
        remarks = """- url :  https://be-univ-dev.teknologioperator.com/snap/v1.1/transfer-va/inquiry
- headers: {
  "Channel-Id": "IB",
  "X-Signature": "mtls_signature_here",
  "X-Timestamp": "2025-07-30T10:00:00+07:00",
  "Authorization": "Bearer token_abc123"
}
- Request Body: {
  "amount": {"value": "10000.00", "currency": "IDR"},
  "virtualAccountNo": "8800123456789",
  "customerNo": "123456"
}
- Response: {"responseCode": "4017101", "responseMessage": "Access Token Invalid"}"""

        result = parse_remarks_bss_hit(remarks)

        assert result is not None
        assert "https://be-univ-dev.teknologioperator.com/snap/v1.1/transfer-va/inquiry" in result['url']
        assert "Channel-Id" in result['headers']
        assert "X-Signature" in result['headers']
        assert "amount" in result['request_body']
        assert "virtualAccountNo" in result['request_body']
        assert "responseCode" in result['response']
        assert "4017101" in result['response']

    def test_real_format_b_via_parse_remarks(self):
        """Test that parse_remarks routes correctly for scenario 7."""
        remarks = """- url :  https://be-univ-dev.teknologioperator.com/snap/v1.1/transfer-va/inquiry
- headers: {"Channel-Id": "IB", "X-Signature": "mtls", "Authorization": "Bearer abc"}
- Request Body: {"virtualAccountNo": "8800123456789"}
- Response: {"responseCode": "4017101", "responseMessage": "Access Token Invalid"}"""

        result = parse_remarks(remarks, "7")

        assert result is not None
        assert "be-univ-dev.teknologioperator.com" in result['url']
        assert "Channel-Id" in result['headers']
        assert "virtualAccountNo" in result['request_body']
        assert "4017101" in result['response']

    def test_format_b_no_bullet_markers(self):
        """Test Format B tanpa bullet markers (content detection only)."""
        remarks = """https://be-univ-dev.teknologioperator.com/snap/v1.1/transfer-va/inquiry
{"Channel-Id": "IB", "X-Signature": "mtls_sig", "X-Timestamp": "2025-07-30T10:00:00+07:00", "Authorization": "Bearer token"}
{"virtualAccountNo": "8800123456789", "amount": {"value": "10000.00", "currency": "IDR"}}
{"responseCode": "4017101", "responseMessage": "Access Token Invalid"}"""

        result = parse_remarks_bss_hit(remarks)

        assert result is not None
        assert "be-univ-dev.teknologioperator.com" in result['url']
        # The JSON with header-like keys should be detected as headers
        assert "Channel-Id" in result['headers'] or "X-Signature" in result['headers']
        assert "responseCode" in result['response']

    def test_format_b_multi_section(self):
        """Test Format B with multiple sections (#validate, #commit)."""
        remarks = """#validate
- url: https://api.example.com/snap/v1.1/transfer-va/inquiry
- headers: {"X-Signature": "sig1", "Authorization": "Bearer t1"}
- Request Body: {"virtualAccountNo": "123"}
- Response: {"responseCode": "2002400", "responseMessage": "Success"}

#commit
- url: https://api.example.com/snap/v1.1/transfer-va/payment
- headers: {"X-Signature": "sig2", "Authorization": "Bearer t2"}
- Request Body: {"virtualAccountNo": "123", "amount": {"value": "50000"}}
- Response: {"responseCode": "2002500", "responseMessage": "Success"}"""

        result = parse_remarks_bss_hit(remarks)

        assert result is not None
        # Should contain data from both sections
        assert "inquiry" in result['url'] or "payment" in result['url']
        assert "VALIDATE" in result['url'] or "COMMIT" in result['url']


class TestContentDetectionHelpers:
    """Test helper functions for content detection."""

    def test_extract_url_from_text(self):
        """Test URL extraction from various line formats."""
        assert _extract_url_from_text("https://api.example.com/v1") == "https://api.example.com/v1"
        assert _extract_url_from_text("URL: https://api.example.com/v1") == "https://api.example.com/v1"
        assert _extract_url_from_text("- url :  https://api.example.com/v1") == "https://api.example.com/v1"
        assert _extract_url_from_text("no url here") == ""

    def test_extract_json_blocks(self):
        """Test JSON block extraction."""
        text = '{"key": "value"} some text {"other": "data"}'
        blocks = _extract_json_blocks(text)
        assert len(blocks) == 2
        assert blocks[0][2] == '{"key": "value"}'
        assert blocks[1][2] == '{"other": "data"}'

    def test_extract_nested_json(self):
        """Test nested JSON block extraction."""
        text = '{"amount": {"value": "10000", "currency": "IDR"}, "key": "val"}'
        blocks = _extract_json_blocks(text)
        assert len(blocks) == 1
        assert "amount" in blocks[0][2]
        assert "currency" in blocks[0][2]

    def test_is_header_like_json(self):
        """Test header JSON detection."""
        header_json = '{"X-Signature": "abc", "Authorization": "Bearer token", "X-Timestamp": "2025-01-01"}'
        assert _is_header_like_json(header_json) is True

        body_json = '{"partnerReferenceNo": "900009", "amount": 50000}'
        assert _is_header_like_json(body_json) is False

    def test_is_response_like_json(self):
        """Test response JSON detection."""
        response_json = '{"responseCode": "2001700", "responseMessage": "Successful"}'
        assert _is_response_like_json(response_json) is True

        body_json = '{"amount": 50000, "destination": "1234567890"}'
        assert _is_response_like_json(body_json) is False

    def test_content_detection_fallback(self):
        """Test full content detection without any labels."""
        text = """https://api.example.com/v1/transfer
{"X-Signature": "sig123", "Authorization": "Bearer token", "X-Timestamp": "2025-01-01"}
{"partnerReferenceNo": "900009", "amount": 50000}
{"responseCode": "2001700", "responseMessage": "Successful"}"""

        result = _parse_by_content_detection(text)

        assert result is not None
        assert "api.example.com" in result['url']
        assert "X-Signature" in result['headers']
        assert "responseCode" in result['response']


class TestCrossParseFallback:
    """Test that parsers fall back to each other when primary parser fails."""

    def test_mitra_format_parsed_by_bss_scenario(self):
        """Test that Mitra format text can still be parsed under BSS scenario."""
        remarks = """URL:
https://api.example.com/v1/transfer

Header:
X-TIMESTAMP: 2025-07-30T14:28:56+07:00
Authorization: Bearer token

RequestBody:
{"amount": 50000}

Response:
{"responseCode":"2001700"}"""

        # Even though scenario 7 normally uses BSS parser,
        # it should fallback to mitra parser
        result = parse_remarks(remarks, "7")

        assert result is not None
        assert "api.example.com" in result['url']

    def test_bss_format_parsed_by_mitra_scenario(self):
        """Test that BSS format text can still be parsed under Mitra scenario."""
        remarks = """- url: https://api.example.com/v1/va
- headers: {"Authorization": "Bearer abc"}
- Request Body: {"data": "test"}
- Response: {"status": "00"}"""

        # Even though scenario 2 normally uses Mitra parser,
        # it should fallback to bss parser
        result = parse_remarks(remarks, "2")

        assert result is not None
        assert "api.example.com" in result['url']

"""
Unit tests untuk fungsi parsing kolom Remarks di main.py.
Menguji kedua format: 'Mitra HIT' dan 'BSS YANG HIT'.
"""

import sys
import os

# Tambahkan root project ke path agar bisa import main
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import (
    parse_remarks_mitra_hit,
    parse_remarks_bss_hit,
    parse_remarks,
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

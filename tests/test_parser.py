"""
Unit tests untuk verifikasi bahwa parser functions sudah dihapus.
Remarks sekarang di-copy langsung ke kolom Request tanpa parsing.
"""

import sys
import os

# Tambahkan root project ke path agar bisa import main
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import main


class TestParserFunctionsRemoved:
    """Verifikasi bahwa semua fungsi parser sudah dihapus dari main.py."""

    def test_parse_remarks_removed(self):
        """Fungsi parse_remarks tidak ada lagi."""
        assert not hasattr(main, 'parse_remarks')

    def test_parse_remarks_mitra_hit_removed(self):
        """Fungsi parse_remarks_mitra_hit tidak ada lagi."""
        assert not hasattr(main, 'parse_remarks_mitra_hit')

    def test_parse_remarks_bss_hit_removed(self):
        """Fungsi parse_remarks_bss_hit tidak ada lagi."""
        assert not hasattr(main, 'parse_remarks_bss_hit')

    def test_parse_with_labels_removed(self):
        """Fungsi _parse_with_labels tidak ada lagi."""
        assert not hasattr(main, '_parse_with_labels')

    def test_parse_by_content_detection_removed(self):
        """Fungsi _parse_by_content_detection tidak ada lagi."""
        assert not hasattr(main, '_parse_by_content_detection')

    def test_parse_bss_single_section_removed(self):
        """Fungsi _parse_bss_single_section tidak ada lagi."""
        assert not hasattr(main, '_parse_bss_single_section')

    def test_parse_bss_multi_section_removed(self):
        """Fungsi _parse_bss_multi_section tidak ada lagi."""
        assert not hasattr(main, '_parse_bss_multi_section')

    def test_extract_url_from_text_removed(self):
        """Fungsi _extract_url_from_text tidak ada lagi."""
        assert not hasattr(main, '_extract_url_from_text')

    def test_extract_json_blocks_removed(self):
        """Fungsi _extract_json_blocks tidak ada lagi."""
        assert not hasattr(main, '_extract_json_blocks')

    def test_is_header_like_json_removed(self):
        """Fungsi _is_header_like_json tidak ada lagi."""
        assert not hasattr(main, '_is_header_like_json')

    def test_is_response_like_json_removed(self):
        """Fungsi _is_response_like_json tidak ada lagi."""
        assert not hasattr(main, '_is_response_like_json')

    def test_is_header_key_value_line_removed(self):
        """Fungsi _is_header_key_value_line tidak ada lagi."""
        assert not hasattr(main, '_is_header_key_value_line')

    def test_bss_hit_scenarios_removed(self):
        """Konstanta BSS_HIT_SCENARIOS tidak ada lagi."""
        assert not hasattr(main, 'BSS_HIT_SCENARIOS')

    def test_mitra_hit_scenarios_removed(self):
        """Konstanta MITRA_HIT_SCENARIOS tidak ada lagi."""
        assert not hasattr(main, 'MITRA_HIT_SCENARIOS')

    def test_header_keys_removed(self):
        """Konstanta _HEADER_KEYS tidak ada lagi."""
        assert not hasattr(main, '_HEADER_KEYS')



# ============================================================
# Test fleksibilitas parser terhadap variasi penanda antar mitra
# ============================================================

from main import split_request_response


class TestParserFleksibel:
    """Parser harus mengenali kata kunci inti (url/header/request/response)
    tanpa terpaku pada kata pengiring, di berbagai variasi format mitra."""

    def _cek_lengkap(self, remarks):
        req, resp = split_request_response(remarks)
        return (
            "URL Endpoint:" in req
            and "Header Request:" in req
            and "Request Body:" in req
            and "Response Body:" in resp
        )

    def test_format_url_headers_request_body_response(self):
        r = ("URL:\nhttps://a.test/x\n\nHeaders:\nAuth: xyz\n\n"
             "Request Body:\n{\"a\":1}\n\nResponse:\n{\"ok\":true}")
        assert self._cek_lengkap(r)

    def test_format_request_url_dan_penanda_menempel(self):
        # Penanda 'Response body:' menempel di akhir baris JSON sebelumnya
        r = ("Request URL: POST https://b.test/y\n"
             "Request headers: {\"H\":[\"v\"]}\n"
             "Request body: {\"b\":2}Response body: {\"ok\":false}")
        assert self._cek_lengkap(r)

    def test_format_url_endpoint_header_request(self):
        r = ("URL Endpoint:\nhttps://c.test/z\n\nHeader Request:\n[Auth: bearer]\n\n"
             "Request Body:\n{\"c\":3}\n\nResponse Body:\n{\"code\":200}")
        assert self._cek_lengkap(r)

    def test_format_url_request_header_saja(self):
        # Mitra pakai "URL Request:" dan "Header:" (tanpa kata 'Request'/'s')
        r = ("URL Request:\nhttps://d.test/w\n\nHeader:\nContent-Type: application/json\n\n"
             "Request Body:\n{\"d\":4}\n\nResponse:\n{\"r\":\"ok\"}")
        assert self._cek_lengkap(r)

    def test_format_payload_sebagai_body(self):
        # 'Payload:' dikenali sebagai Request Body
        r = ("URL:\nhttps://e.test\n\nHeader:\nX: y\n\nPayload:\n{\"e\":5}\n\n"
             "Response:\n{\"z\":9}")
        assert self._cek_lengkap(r)



# ============================================================
# Test format header: tidak boleh menghasilkan kutip ganda ("")
# ============================================================

from main import _format_headers


class TestFormatHeaderKutipGanda:
    """Header value dengan karakter base64 (+ / =) atau input yang sudah
    berupa array 'Key=Value' TIDAK boleh menghasilkan kutip ganda."""

    TOKEN = "Bearer 8KGjr+yYgfCqqqHtgo/slqnpoa/wpYW/8KKylOuQvfCikYPonqc="

    def test_input_sudah_array_key_value(self):
        h = '[\n  "Content-Type=application/json",\n  "Authorization=%s"\n]' % self.TOKEN
        out = _format_headers(h)
        assert '""' not in out
        assert 'Authorization=%s' % self.TOKEN in out

    def test_input_json_object_string(self):
        h = '{"Authorization": "%s", "Content-Type": "application/json"}' % self.TOKEN
        out = _format_headers(h)
        assert '""' not in out
        assert "Authorization=%s" % self.TOKEN in out

    def test_input_json_object_array_value(self):
        h = '{"Authorization": ["%s"]}' % self.TOKEN
        out = _format_headers(h)
        assert '""' not in out
        assert "Authorization=%s" % self.TOKEN in out

    def test_input_baris_polos(self):
        h = "[\nContent-Type: application/json\nAuthorization: %s\n]" % self.TOKEN
        out = _format_headers(h)
        assert '""' not in out
        assert "Authorization=%s" % self.TOKEN in out

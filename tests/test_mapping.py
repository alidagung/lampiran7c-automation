"""
Unit tests untuk fungsi mapping skenario di main.py.
Menguji logika "fill empty only" untuk section yang di-share.
"""

import sys
import os

# Tambahkan root project ke path agar bisa import main
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import (
    extract_sub_number,
    map_uat_to_lampiran,
    LAMPIRAN_SECTIONS,
)


class TestExtractSubNumber:
    """Test fungsi extract_sub_number."""

    def test_standard_format(self):
        """Test format standar nomor kasus."""
        assert extract_sub_number("1.1") == 1
        assert extract_sub_number("2.3") == 3
        assert extract_sub_number("7.15") == 15
        assert extract_sub_number("9.7") == 7

    def test_double_digit(self):
        """Test dengan sub-number dua digit."""
        assert extract_sub_number("7.33") == 33
        assert extract_sub_number("6.13") == 13

    def test_invalid_format(self):
        """Test dengan format yang tidak valid."""
        assert extract_sub_number("") is None
        assert extract_sub_number("abc") is None
        assert extract_sub_number("1") is None

    def test_non_numeric_sub(self):
        """Test dengan sub-number non-numerik."""
        assert extract_sub_number("1.abc") is None


class TestMapUatToLampiran:
    """Test fungsi map_uat_to_lampiran."""

    def _make_row(self, nomor_kasus, remarks="", hasil_aktual="Berhasil", langkah_tes="", hasil_diharapkan=""):
        """Helper untuk membuat row data UAT."""
        return {
            'nomor_skenario': nomor_kasus.split('.')[0] if '.' in nomor_kasus else "",
            'nomor_kasus_tes': nomor_kasus,
            'langkah_tes': langkah_tes or f"Test case {nomor_kasus}",
            'hasil_diharapkan': hasil_diharapkan or "Berhasil",
            'hasil_aktual': hasil_aktual,
            'remarks': remarks,
        }

    def test_basic_mapping_balance(self):
        """Test mapping sederhana untuk Balance Services."""
        uat_data = {
            "Balance Services": [
                self._make_row("1.1", remarks="URL:\nhttps://api.test.com/balance\n\nResponse:\n{\"ok\":true}"),
                self._make_row("1.2", remarks="URL:\nhttps://api.test.com/balance2\n\nResponse:\n{\"ok\":false}"),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        assert result["API Balance Inquiry"][0] is not None
        assert result["API Balance Inquiry"][1] is not None
        assert result["API Balance Inquiry"][2] is None  # Tidak ada data
        assert result["API Balance Inquiry"][0]['no'] == 1
        assert result["API Balance Inquiry"][1]['no'] == 2

    def test_basic_mapping_intrabank(self):
        """Test mapping untuk Intrabank Transfer."""
        uat_data = {
            "Intrabank Transfer": [
                self._make_row("2.1", remarks="URL:\nhttps://api.test.com/intra\n\nHeader:\n{\"auth\":\"x\"}\n\nRequestBody:\n{\"amt\":100}\n\nResponse:\n{\"ok\":true}"),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        assert result["Intrabank Transfer"][0] is not None
        assert "URL Endpoint:" in result["Intrabank Transfer"][0]['request']
        assert "Response Body:" in result["Intrabank Transfer"][0]['response']

    def test_fill_empty_only_interbank(self):
        """
        Test logika 'fill empty only' untuk Interbank Transfer.
        Skenario 3 mengisi terlebih dahulu, Skenario 4 hanya mengisi yang kosong.
        """
        uat_data = {
            "Interbank Transfer": [
                self._make_row("3.1", remarks="URL:\nhttps://api.test.com/interbank/1\n\nResponse:\n{\"r\":1}"),
                self._make_row("3.3", remarks="URL:\nhttps://api.test.com/interbank/3\n\nResponse:\n{\"r\":3}"),
                # 3.2 tidak ada datanya
            ],
            "Interbank Transfer via BI FAST": [
                self._make_row("4.1", remarks="URL:\nhttps://api.test.com/bifast/1\n\nResponse:\n{\"r\":\"bifast1\"}"),
                self._make_row("4.2", remarks="URL:\nhttps://api.test.com/bifast/2\n\nResponse:\n{\"r\":\"bifast2\"}"),
                self._make_row("4.3", remarks="URL:\nhttps://api.test.com/bifast/3\n\nResponse:\n{\"r\":\"bifast3\"}"),
            ],
        }

        result = map_uat_to_lampiran(uat_data)

        # Row 1 (index 0): diisi oleh skenario 3 -> URL harus interbank/1
        assert result["Interbank Transfer"][0] is not None
        assert "interbank/1" in result["Interbank Transfer"][0]['request']

        # Row 2 (index 1): TIDAK ada di skenario 3, diisi oleh skenario 4 -> URL harus bifast/2
        assert result["Interbank Transfer"][1] is not None
        assert "bifast/2" in result["Interbank Transfer"][1]['request']

        # Row 3 (index 2): sudah diisi skenario 3 -> skenario 4 TIDAK overwrite
        assert result["Interbank Transfer"][2] is not None
        assert "interbank/3" in result["Interbank Transfer"][2]['request']
        # Pastikan BUKAN bifast/3
        assert "bifast/3" not in result["Interbank Transfer"][2]['request']

    def test_fill_empty_only_virtual_account(self):
        """
        Test logika 'fill empty only' untuk Virtual Account.
        Skenario 7 mengisi terlebih dahulu, Skenario 8 dan 9 hanya mengisi yang kosong.
        """
        uat_data = {
            "Transfer VA": [
                self._make_row("7.1", remarks="- url: https://api.test.com/va/1\n- headers: {}\n- body: {}\n- response: {\"r\":\"va1\"}"),
                self._make_row("7.3", remarks="- url: https://api.test.com/va/3\n- headers: {}\n- body: {}\n- response: {\"r\":\"va3\"}"),
            ],
            "Transfer VA Prima": [
                self._make_row("8.1", remarks="- url: https://api.test.com/prima/1\n- headers: {}\n- body: {}\n- response: {\"r\":\"prima1\"}"),
                self._make_row("8.2", remarks="- url: https://api.test.com/prima/2\n- headers: {}\n- body: {}\n- response: {\"r\":\"prima2\"}"),
            ],
            "Transfer VA BI FAST": [
                self._make_row("9.1", remarks="- url: https://api.test.com/vabifast/1\n- headers: {}\n- body: {}\n- response: {\"r\":\"vabifast1\"}"),
                self._make_row("9.2", remarks="- url: https://api.test.com/vabifast/2\n- headers: {}\n- body: {}\n- response: {\"r\":\"vabifast2\"}"),
                self._make_row("9.3", remarks="- url: https://api.test.com/vabifast/3\n- headers: {}\n- body: {}\n- response: {\"r\":\"vabifast3\"}"),
            ],
        }

        result = map_uat_to_lampiran(uat_data)

        # Row 1 (index 0): diisi oleh skenario 7 -> URL harus va/1
        assert result["API Virtual Account"][0] is not None
        assert "va/1" in result["API Virtual Account"][0]['request']
        # Pastikan BUKAN prima/1 atau vabifast/1
        assert "prima/1" not in result["API Virtual Account"][0]['request']
        assert "vabifast/1" not in result["API Virtual Account"][0]['request']

        # Row 2 (index 1): TIDAK ada di skenario 7, diisi oleh skenario 8 -> prima/2
        assert result["API Virtual Account"][1] is not None
        assert "prima/2" in result["API Virtual Account"][1]['request']

        # Row 3 (index 2): sudah diisi skenario 7 -> skenario 8 dan 9 TIDAK overwrite
        assert result["API Virtual Account"][2] is not None
        assert "va/3" in result["API Virtual Account"][2]['request']
        assert "vabifast/3" not in result["API Virtual Account"][2]['request']

    def test_fill_empty_va_priority_8_before_9(self):
        """
        Test bahwa skenario 8 diproses sebelum skenario 9 untuk VA.
        Jika row kosong di skenario 7, skenario 8 mengisi duluan.
        """
        uat_data = {
            "Transfer VA": [
                # Hanya row 1, sisanya kosong
                self._make_row("7.1", remarks="- url: https://api.test.com/va/1\n- headers: {}\n- body: {}\n- response: {}"),
            ],
            "Transfer VA Prima": [
                self._make_row("8.2", remarks="- url: https://api.test.com/prima/2\n- headers: {}\n- body: {}\n- response: {}"),
                self._make_row("8.4", remarks="- url: https://api.test.com/prima/4\n- headers: {}\n- body: {}\n- response: {}"),
            ],
            "Transfer VA BI FAST": [
                self._make_row("9.2", remarks="- url: https://api.test.com/vabifast/2\n- headers: {}\n- body: {}\n- response: {}"),
                self._make_row("9.3", remarks="- url: https://api.test.com/vabifast/3\n- headers: {}\n- body: {}\n- response: {}"),
                self._make_row("9.4", remarks="- url: https://api.test.com/vabifast/4\n- headers: {}\n- body: {}\n- response: {}"),
            ],
        }

        result = map_uat_to_lampiran(uat_data)

        # Row 1 (index 0): diisi skenario 7
        assert result["API Virtual Account"][0] is not None
        assert "va/1" in result["API Virtual Account"][0]['request']

        # Row 2 (index 1): kosong di 7, diisi skenario 8 (prima/2)
        assert result["API Virtual Account"][1] is not None
        assert "prima/2" in result["API Virtual Account"][1]['request']

        # Row 3 (index 2): kosong di 7, tidak ada di 8, diisi skenario 9 (vabifast/3)
        assert result["API Virtual Account"][2] is not None
        assert "vabifast/3" in result["API Virtual Account"][2]['request']

        # Row 4 (index 3): kosong di 7, diisi skenario 8 (prima/4), skenario 9 TIDAK overwrite
        assert result["API Virtual Account"][3] is not None
        assert "prima/4" in result["API Virtual Account"][3]['request']
        assert "vabifast/4" not in result["API Virtual Account"][3]['request']

    def test_tidak_dites_handling(self):
        """Test handling ketika Hasil Aktual = 'Tidak dites'."""
        uat_data = {
            "Balance Services": [
                self._make_row("1.1", hasil_aktual="Tidak dites", remarks=""),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        assert result["API Balance Inquiry"][0] is not None
        assert result["API Balance Inquiry"][0]['request'] == ""
        assert result["API Balance Inquiry"][0]['response'] == ""
        assert result["API Balance Inquiry"][0]['notes'] == "Tidak dites"
        assert result["API Balance Inquiry"][0]['result'] == "Tidak dites"

    def test_berhasil_without_remarks(self):
        """Test handling ketika Berhasil tapi Remarks kosong."""
        uat_data = {
            "Balance Services": [
                self._make_row("1.1", hasil_aktual="Berhasil", remarks=""),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        assert result["API Balance Inquiry"][0] is not None
        assert result["API Balance Inquiry"][0]['request'] == ""
        assert result["API Balance Inquiry"][0]['response'] == ""

    def test_empty_uat_data(self):
        """Test dengan data UAT kosong."""
        result = map_uat_to_lampiran({})

        # Semua section harus ada tapi kosong (filled with None)
        for section_name, count in LAMPIRAN_SECTIONS:
            assert section_name in result
            assert len(result[section_name]) == count
            assert all(r is None for r in result[section_name])

    def test_out_of_range_sub_number(self):
        """Test bahwa nomor kasus di luar range diabaikan."""
        uat_data = {
            "Balance Services": [
                self._make_row("1.1", remarks="URL:\nhttps://test.com\n\nResponse:\n{}"),
                self._make_row("1.99", remarks="URL:\nhttps://test.com/out\n\nResponse:\n{}"),  # Out of range (max 11)
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        assert result["API Balance Inquiry"][0] is not None  # 1.1 valid
        # 1.99 should be ignored since Balance only has 11 rows
        assert all(r is None for r in result["API Balance Inquiry"][1:])  # Only row 1 filled

    def test_sknbi_mapping(self):
        """Test mapping untuk SKNBI Transfer (Skenario 6)."""
        uat_data = {
            "SKNBI Transfer": [
                self._make_row("6.1", remarks="URL:\nhttps://api.test.com/sknbi\n\nHeader:\n{\"auth\":\"x\"}\n\nRequestBody:\n{\"amt\":200}\n\nResponse:\n{\"ok\":true}"),
                self._make_row("6.5", remarks="URL:\nhttps://api.test.com/sknbi5\n\nResponse:\n{\"ok\":true}"),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        assert result["API SKNBI Transfer"][0] is not None
        assert result["API SKNBI Transfer"][4] is not None
        assert "sknbi" in result["API SKNBI Transfer"][0]['request']
        assert "sknbi5" in result["API SKNBI Transfer"][4]['request']

    def test_section_initialization(self):
        """Test bahwa semua section diinisialisasi dengan jumlah row yang benar."""
        result = map_uat_to_lampiran({})

        assert len(result["API Balance Inquiry"]) == 11
        assert len(result["Intrabank Transfer"]) == 12
        assert len(result["Interbank Transfer"]) == 13
        assert len(result["API RTGS Transfer"]) == 13
        assert len(result["API SKNBI Transfer"]) == 13
        assert len(result["API Virtual Account"]) == 33

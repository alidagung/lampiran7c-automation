"""
Unit tests untuk fungsi mapping skenario di main.py.
Menguji logika "fill empty only" untuk section yang di-share.
Menguji bahwa Remarks di-copy langsung ke kolom Request tanpa parsing.
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
            'nama_modul': '',
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
                self._make_row("1.1", remarks="some remarks content"),
                self._make_row("1.2", remarks="other remarks content"),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        assert result["API Balance Inquiry"][0] is not None
        assert result["API Balance Inquiry"][1] is not None
        assert result["API Balance Inquiry"][2] is None  # Tidak ada data
        assert result["API Balance Inquiry"][0]['no'] == 1
        assert result["API Balance Inquiry"][1]['no'] == 2

    def test_remarks_copied_directly_to_request(self):
        """Test bahwa Remarks dipisah & diformat: Request berisi URL/Header/Body,
        Response berisi bagian setelah penanda Response."""
        remarks_text = """URL:
https://api.example.com/v1/transfer

Headers:
X-TIMESTAMP: 2025-07-30T14:28:56+07:00
Authorization: Bearer token123

Request Body:
{"amount": 50000}

Response:
{"responseCode":"2001700","responseMessage":"Successful"}"""

        uat_data = {
            "Balance Services": [
                self._make_row("1.1", remarks=remarks_text, hasil_aktual="Berhasil"),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        row = result["API Balance Inquiry"][0]
        assert row is not None
        # Request memuat URL Endpoint, Header Request, dan Request Body (diformat)
        assert "URL Endpoint:" in row['request']
        assert "https://api.example.com/v1/transfer" in row['request']
        assert "Header Request:" in row['request']
        assert '"Authorization=Bearer token123"' in row['request']
        assert "Request Body:" in row['request']
        # Response body dipindah ke kolom Response, tidak ada di Request
        assert "Response Body:" in row['response']
        assert "2001700" in row['response']
        assert "2001700" not in row['request']

    def test_remarks_unknown_format_fallback(self):
        """Test fallback: jika format Remarks tidak dikenali (tanpa penanda
        URL/Headers/Request Body/Response standar), data tetap tidak hilang."""
        remarks_text = "catatan bebas tanpa struktur baku"

        uat_data = {
            "Transfer VA": [
                self._make_row("7.1", remarks=remarks_text, hasil_aktual="Berhasil"),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        row = result["API Virtual Account"][0]
        assert row is not None
        # Data mentah tetap dipertahankan di Request agar tidak hilang
        assert row['request'] == remarks_text
        assert row['response'] == ""

    def test_basic_mapping_intrabank(self):
        """Test mapping untuk Intrabank Transfer - Remarks dipisah & diformat."""
        remarks_text = "URL:\nhttps://api.test.com/intra\n\nHeaders:\nauth: x\n\nRequest Body:\n{\"amt\":100}\n\nResponse:\n{\"ok\":true}"
        uat_data = {
            "Intrabank Transfer": [
                self._make_row("2.1", remarks=remarks_text),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        row = result["Intrabank Transfer"][0]
        assert row is not None
        # Request memuat URL Endpoint + Request Body, Response terpisah
        assert "URL Endpoint:" in row['request']
        assert "https://api.test.com/intra" in row['request']
        assert "Request Body:" in row['request']
        assert "Response Body:" in row['response']
        assert '"ok":true' in row['response'] or '"ok": true' in row['response']

    def test_interbank_bifast_not_mapped(self):
        """
        Aturan bisnis: 'Interbank Transfer via BI FAST' (skenario 4.x) TIDAK
        dipindahkan ke Lampiran 7C. Data BI FAST tidak boleh mengisi baris
        kosong di section Interbank Transfer.
        """
        uat_data = {
            "Interbank Transfer": [
                self._make_row("3.1", remarks="interbank remarks 1"),
                self._make_row("3.3", remarks="interbank remarks 3"),
                # 3.2 sengaja tidak ada datanya
            ],
            "Interbank Transfer via BI FAST": [
                self._make_row("4.1", remarks="bifast remarks 1"),
                self._make_row("4.2", remarks="bifast remarks 2"),
                self._make_row("4.3", remarks="bifast remarks 3"),
            ],
        }

        result = map_uat_to_lampiran(uat_data)

        # Row 1 (index 0): diisi oleh skenario 3
        assert result["Interbank Transfer"][0] is not None
        assert result["Interbank Transfer"][0]['request'] == "interbank remarks 1"

        # Row 2 (index 1): 3.2 tidak ada -> TETAP KOSONG (BI FAST tidak mengisi)
        assert result["Interbank Transfer"][1] is None

        # Row 3 (index 2): diisi skenario 3
        assert result["Interbank Transfer"][2] is not None
        assert result["Interbank Transfer"][2]['request'] == "interbank remarks 3"

        # Tidak ada satupun data BI FAST yang bocor ke Interbank Transfer
        for row in result["Interbank Transfer"]:
            if row is not None:
                assert "bifast" not in row['request']

    def test_fill_empty_only_virtual_account(self):
        """
        Aturan bisnis: 'Transfer VA Prima' (8.x) dan 'Transfer VA BI FAST' (9.x)
        TIDAK dipindahkan ke Lampiran 7C. Hanya 'Transfer VA' (7.x) yang masuk.
        """
        uat_data = {
            "Transfer VA": [
                self._make_row("7.1", remarks="va remarks 1"),
                self._make_row("7.3", remarks="va remarks 3"),
                # 7.2 sengaja tidak ada datanya
            ],
            "Transfer VA Prima": [
                self._make_row("8.1", remarks="prima remarks 1"),
                self._make_row("8.2", remarks="prima remarks 2"),
            ],
            "Transfer VA BI FAST": [
                self._make_row("9.1", remarks="vabifast remarks 1"),
                self._make_row("9.2", remarks="vabifast remarks 2"),
            ],
        }

        result = map_uat_to_lampiran(uat_data)

        # Row 1 (index 0): diisi oleh skenario 7
        assert result["API Virtual Account"][0] is not None
        assert result["API Virtual Account"][0]['request'] == "va remarks 1"

        # Row 2 (index 1): 7.2 tidak ada -> TETAP KOSONG (Prima/BI FAST tidak mengisi)
        assert result["API Virtual Account"][1] is None

        # Row 3 (index 2): diisi skenario 7
        assert result["API Virtual Account"][2] is not None
        assert result["API Virtual Account"][2]['request'] == "va remarks 3"

        # Tidak ada satupun data Prima / BI FAST yang bocor ke Virtual Account
        for row in result["API Virtual Account"]:
            if row is not None:
                assert "prima" not in row['request']
                assert "vabifast" not in row['request']

    def test_va_prima_bifast_tidak_dipindahkan(self):
        """Jika HANYA ada data VA Prima & BI FAST (tanpa VA biasa), maka
        section Virtual Account harus KOSONG seluruhnya."""
        uat_data = {
            "Transfer VA Prima": [self._make_row("8.1", remarks="prima 1")],
            "Transfer VA BI FAST": [self._make_row("9.1", remarks="bifast 1")],
        }
        result = map_uat_to_lampiran(uat_data)
        assert all(r is None for r in result["API Virtual Account"])

    def test_tidak_dites_handling(self):
        """Test handling ketika Hasil Aktual = 'Tidak dites'.
        Baris tetap ditampilkan (Result=N/A), tapi Notes dikosongkan."""
        uat_data = {
            "Balance Services": [
                self._make_row("1.1", hasil_aktual="Tidak dites", remarks="Catatan dari remarks"),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        assert result["API Balance Inquiry"][0] is not None
        assert result["API Balance Inquiry"][0]['request'] == ""
        assert result["API Balance Inquiry"][0]['response'] == ""
        # Remark dipindah ke Notes untuk baris "Tidak dites"
        assert result["API Balance Inquiry"][0]['notes'] == "Catatan dari remarks"
        # Result mapped: "Tidak dites" -> "N/A"
        assert result["API Balance Inquiry"][0]['result'] == "N/A"

    def test_tidak_dites_empty_remarks(self):
        """Test handling ketika Hasil Aktual = 'Tidak dites' dan Remarks kosong.
        Notes tetap kosong karena tidak ada Remark."""
        uat_data = {
            "Balance Services": [
                self._make_row("1.1", hasil_aktual="Tidak dites", remarks=""),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        assert result["API Balance Inquiry"][0] is not None
        assert result["API Balance Inquiry"][0]['request'] == ""
        assert result["API Balance Inquiry"][0]['response'] == ""
        assert result["API Balance Inquiry"][0]['notes'] == ""

    def test_belum_dites_handling(self):
        """Test handling ketika Hasil Aktual = 'Belum dites'.
        Sama seperti 'Tidak dites': Result=N/A, Request/Response kosong,
        dan Remark dipindah ke Notes."""
        uat_data = {
            "Balance Services": [
                self._make_row("1.1", hasil_aktual="Belum dites", remarks="apapun isinya"),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        row = result["API Balance Inquiry"][0]
        assert row is not None
        assert row['request'] == ""
        assert row['response'] == ""
        assert row['notes'] == "apapun isinya"
        assert row['result'] == "N/A"

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

    def test_berhasil_with_remarks(self):
        """Test handling ketika Berhasil dan Remarks ada isinya."""
        uat_data = {
            "Balance Services": [
                self._make_row("1.1", hasil_aktual="Berhasil", remarks="Full remarks content here"),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        assert result["API Balance Inquiry"][0] is not None
        # Seluruh Remarks masuk ke Request
        assert result["API Balance Inquiry"][0]['request'] == "Full remarks content here"
        # Response kosong
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
                self._make_row("1.1", remarks="valid row"),
                self._make_row("1.99", remarks="out of range"),  # Out of range (max 11)
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
                self._make_row("6.1", remarks="sknbi remarks 1"),
                self._make_row("6.5", remarks="sknbi remarks 5"),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        assert result["API SKNBI Transfer"][0] is not None
        assert result["API SKNBI Transfer"][4] is not None
        assert result["API SKNBI Transfer"][0]['request'] == "sknbi remarks 1"
        assert result["API SKNBI Transfer"][4]['request'] == "sknbi remarks 5"

    def test_section_initialization(self):
        """Test bahwa semua section diinisialisasi dengan jumlah row yang benar."""
        result = map_uat_to_lampiran({})

        assert len(result["API Balance Inquiry"]) == 11
        assert len(result["Intrabank Transfer"]) == 12
        assert len(result["Interbank Transfer"]) == 13
        assert len(result["API RTGS Transfer"]) == 13
        assert len(result["API SKNBI Transfer"]) == 13
        assert len(result["API Virtual Account"]) == 33

    def test_response_extracted_to_response_column(self):
        """Test bahwa bagian setelah penanda Response dipindah ke kolom Response,
        dan tidak lagi ikut di kolom Request."""
        remarks_with_response = """URL:
https://api.test.com/endpoint

Response:
{"responseCode": "200", "responseMessage": "Success"}"""

        uat_data = {
            "Balance Services": [
                self._make_row("1.1", remarks=remarks_with_response, hasil_aktual="Berhasil"),
            ]
        }

        result = map_uat_to_lampiran(uat_data)

        row = result["API Balance Inquiry"][0]
        assert row is not None
        # Bagian Response dipindah ke kolom Response
        assert "Response Body:" in row['response']
        assert '"responseCode"' in row['response']
        # URL tetap di Request, tapi isi Response TIDAK bocor ke Request
        assert "URL Endpoint:" in row['request']
        assert "Success" not in row['request']

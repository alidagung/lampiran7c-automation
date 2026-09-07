"""
Unit tests untuk deteksi produk (detect_product) & profil produk.
Memastikan tool memilih konfigurasi yang tepat: Fund Transfer/VA vs QRIS.
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import detect_product, PRODUCT_PROFILES


class TestDetectProduct:
    def test_qris_dari_qr_mpm(self):
        uat = {"Balance Services": [], "QR MPM": [],
               "API Transaction History List": []}
        assert detect_product(uat) == "QRIS"

    def test_ft_va_diprioritaskan_meski_ada_pengecekan_mutasi(self):
        # File FT/VA umum juga punya "PENGECEKAN MUTASI DAN JURNAL".
        # Penanda FT/VA harus menang -> tidak boleh salah dikira QRIS.
        uat = {
            "Balance Services": [],
            "Transfer VA": [],
            "RTGS Transfer": [],
            "PENGECEKAN MUTASI DAN JURNAL": [],
        }
        assert detect_product(uat) == "FT_VA"

    def test_pengecekan_mutasi_saja_bukan_penanda_qris(self):
        # Hanya ada Balance + Pengecekan Mutasi (tanpa QR MPM) -> default FT_VA
        uat = {"Balance Services": [], "PENGECEKAN MUTASI DAN JURNAL": []}
        assert detect_product(uat) == "FT_VA"

    def test_ft_va_penanda_klasik(self):
        for marker in ["Transfer VA", "RTGS Transfer", "SKNBI Transfer",
                       "Interbank Transfer", "Intrabank Transfer"]:
            assert detect_product({marker: []}) == "FT_VA", marker

    def test_default_ft_va_saat_tak_dikenali(self):
        assert detect_product({"Sesuatu Lain": []}) == "FT_VA"


class TestProfilQRIS:
    def test_qris_hanya_3_section(self):
        prof = PRODUCT_PROFILES["QRIS"]
        names = [n for n, _ in prof["sections"]]
        assert names == ["Balance Inquiry", "API Transaction History List", "QR MPM"]

    def test_qris_tidak_renumber(self):
        # QRIS pakai nomor kasus ASLI (tidak dinomori ulang)
        assert PRODUCT_PROFILES["QRIS"]["renumber"] is False

    def test_qris_skip_pengecekan_mutasi(self):
        # Tidak ada mapping untuk Pengecekan Mutasi
        targets = [t for _, _, t, _ in PRODUCT_PROFILES["QRIS"]["mapping"]]
        assert "PENGECEKAN MUTASI DAN JURNAL" not in targets

    def test_ft_va_renumber(self):
        assert PRODUCT_PROFILES["FT_VA"]["renumber"] is True

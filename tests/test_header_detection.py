"""
Regresi: header row TABEL tidak boleh tertukar dengan blok metadata di atas
tabel yang juga memakai kata "Kategori" (mis. "Kategori | Simulasi").

Bug asal: file dengan judul dokumen mengandung "Virtual Account" + blok
metadata "Kategori | Simulasi" membuat parser salah mulai membaca terlalu awal,
sehingga judul dokumen terdeteksi sebagai section "Transfer VA" dan baris
tanda tangan ("Irdian ... | Tanda Tangan") bocor menjadi case.
"""

import sys
import os
import io

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

from main import read_uat_script


def _make_uat_with_metadata():
    """
    Bangun workbook meniru struktur file mitra: ada judul dokumen mengandung
    'Virtual Account', blok metadata (Kategori|Simulasi, Nomor Referensi + nama
    penandatangan), lalu header tabel sebenarnya, lalu data.
    Kolom: A kosong, B=Kategori, C=Nama Modul, E=Nomor Kasus, F=Langkah Tes,
    H=Hasil Aktual, I=Remarks (0-based index: 1,2,4,5,7,8).
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "UAT Script"

    def setrow(rownum, mapping):
        for col_idx0, val in mapping.items():
            ws.cell(row=rownum, column=col_idx0 + 1, value=val)

    # Baris judul (mengandung "Virtual Account") -> jangan jadi section
    setrow(2, {1: "User Acceptance Test Script Penambahan layanan Virtual Account Fund Transfer"})
    # Blok metadata dgn kata "Kategori" -> jangan jadi header row
    setrow(6, {1: "Nomor Referensi", 2: "SM-XX", 4: "Nur Cahya", 5: "Tanda Tangan", 7: "Wibiyanto"})
    setrow(7, {1: "Kategori", 2: "Simulasi"})
    setrow(8, {1: "Lingkungan Tes", 2: "UAT"})
    # Baris tanda tangan (nama di kolom E, "Tanda Tangan" di kolom F)
    setrow(14, {1: "Nomor Referensi", 2: "SM-XX", 4: "Irdian Muhammad Haikal", 5: "Tanda Tangan", 7: "Dodi"})
    # HEADER TABEL SEBENARNYA
    setrow(21, {1: "Kategori", 2: "Nama Modul", 4: "Nomor Kasus Tes", 5: "Langkah Tes", 7: "Hasil Aktual"})
    # Section + data
    setrow(22, {1: "Balance Services"})
    setrow(23, {1: "Input Data", 2: "Any Service", 4: "1.1", 5: "Access Token Invalid", 7: "Berhasil", 8: "URL:\nhttps://x"})
    setrow(24, {1: "Transfer VA"})
    setrow(25, {1: "Input Data", 2: "VA", 4: "7.1", 5: "Access Token Invalid", 7: "Berhasil", 8: "URL:\nhttps://y"})

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


class TestHeaderRowDetection:
    def test_section_pertama_balance_bukan_va(self):
        data = read_uat_script(_make_uat_with_metadata())
        sections = list(data.keys())
        assert sections[0] == "Balance Services", sections
        assert "Transfer VA" in sections

    def test_baris_tanda_tangan_tidak_bocor(self):
        data = read_uat_script(_make_uat_with_metadata())
        for sec, rows in data.items():
            for r in rows:
                blob = (str(r["nomor_kasus_tes"]) + " " + str(r["langkah_tes"])).lower()
                assert "irdian" not in blob
                assert "tanda tangan" not in blob
                assert "nomor kasus" not in blob

    def test_judul_dokumen_tidak_jadi_section(self):
        data = read_uat_script(_make_uat_with_metadata())
        # Hanya 2 section valid yang terbaca
        assert set(data.keys()) == {"Balance Services", "Transfer VA"}

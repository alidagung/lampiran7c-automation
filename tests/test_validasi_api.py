"""
Validasi Sub API dari UAT Script (Nama Modul + status) dan penyaringan baris
Sub API pada 5B/6B dengan sel gabungan yang tetap rapi.
"""
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document
from docx.oxml.ns import qn

import lampiran5 as L5
import lampiran6 as L6


def _case(no, modul, hasil):
    return {"nomor_kasus_tes": no, "nama_modul": modul, "hasil_aktual": hasil}


UAT = {
    "Balance Services": [_case("1.1", "Any Service", "Berhasil"), _case("1.6", "Balance Inquiry", "Berhasil")],
    "Intrabank Transfer": [
        _case("2.6", "Internal Account Inquiry", "Tidak dites"),
        _case("2.9", "Intrabank Trigger Transfer", "Berhasil"),
        _case("2.12", "Transaction Status Inquiry", "Berhasil"),
    ],
    "RTGS Transfer": [_case("5.10", "RTGS Trigger Transfer", "Tidak dites")],
    "Transfer VA": [
        _case("7.6", "Inquiry", "Tidak dites"),
        _case("7.10", "Payment", "Berhasil"),
        _case("7.13", "Inquiry Status VA", "Belum dites"),
    ],
    "SKNBI Transfer": [_case("6.10", "Modul Lain", "Berhasil")],
}


def _per_sub(hasil):
    return {v["sub_api"]: v for v in hasil}


class TestValidasiApi:
    def test_status_per_sub_api(self):
        v = _per_sub(L5.validasi_api(UAT))
        assert v["API Balance Inquiry"]["status"] == L5.STATUS_DIGUNAKAN
        assert v["API Balance Inquiry"]["bukti"] == "1.6 Berhasil"
        assert v["API Internal Account Inquiry"]["status"] == L5.STATUS_TIDAK
        assert v["API Intrabank Transfer"]["digunakan"] is True
        assert v["API RTGS Transfer"]["status"] == L5.STATUS_TIDAK
        assert v["API Virtual Account – Inquiry"]["digunakan"] is False
        assert v["API Virtual Account – Payment"]["digunakan"] is True
        assert v["API Virtual Account – Inquiry Status"]["digunakan"] is True   # Belum dites = dipakai
        assert v[L5.SUB_API_TSI]["bukti"].startswith("2.12")

    def test_any_service_tidak_dihitung(self):
        # Balance hanya punya Any Service -> modul Balance Inquiry tidak ada
        uat = {"Balance Services": [_case("1.1", "Any Service", "Berhasil")]}
        v = _per_sub(L5.validasi_api(uat))["API Balance Inquiry"]
        assert v["status"] == L5.STATUS_CEK and v["digunakan"] is True   # diperkirakan dari section

    def test_modul_tidak_dikenali_perlu_dicek(self):
        v = _per_sub(L5.validasi_api(UAT))["API SKNBI Transfer"]
        assert v["status"] == L5.STATUS_CEK and v["digunakan"] is True

    def test_section_tidak_ada(self):
        v = _per_sub(L5.validasi_api(UAT))["API External Account Inquiry"]
        assert v["status"] == L5.STATUS_TIDAK and v["digunakan"] is False

    def test_tsi_tanpa_modul_tidak_digunakan(self):
        uat = {"Intrabank Transfer": [_case("2.9", "Intrabank Trigger Transfer", "Berhasil")]}
        v = _per_sub(L5.validasi_api(uat))[L5.SUB_API_TSI]
        assert v["digunakan"] is False and v["status"] == L5.STATUS_CEK

    def test_layanan_digunakan_turunan_sub_api(self):
        assert L5.layanan_digunakan(UAT) == [
            "API Balance Inquiry", "API Intrabank Transfer", "API SKNBI Transfer",
            "API Virtual Account", L5.LAYANAN_TSI]


def _baris(doc, n_kolom=5):
    out = []
    for tr in doc.tables[0]._tbl.tr_lst[L5.BARIS_DATA_MULAI:]:
        out.append([(L5._vmerge(tc), "".join(x.text or "" for x in tc.iter(qn("w:t"))))
                    for tc in tr.tc_lst[:n_kolom]])
    return out


class TestSaringSubApi:
    PILIH = {"API Balance Inquiry", "API Intrabank Transfer",
             "API Virtual Account – Payment", "API Virtual Account – Inquiry Status"}

    def test_5b_baris_pertama_grup_terhapus_isi_pindah(self):
        rows = _baris(L5.build_lampiran_5b("SM-1", "bank@x", sub_api_tampil=self.PILIH))
        assert [r[2][1] for r in rows] == ["API Balance Inquiry", "API Intrabank Transfer",
                                          "API Virtual Account – Payment",
                                          "API Virtual Account – Inquiry Status"]
        # Intrabank tinggal 1 baris -> tanpa gabungan, isi lengkap
        assert rows[1][0] == (None, "2") and rows[1][1] == (None, "API Intrabank Transfer")
        assert rows[1][3] == (None, "bank@x")
        # VA 2 baris -> restart + continue, isi di baris pertama
        assert rows[2][0] == ("restart", "3") and rows[2][1] == ("restart", "API Virtual Account")
        assert rows[2][3] == ("restart", "bank@x")
        assert rows[3][0] == ("continue", "") and rows[3][1] == ("continue", "")

    def test_6b_kolom_pengguna_dan_user_ikut_pindah(self):
        rows = _baris(L6.build_lampiran_6b("PT Mitra", "m@x", sub_api_tampil=self.PILIH))
        assert rows[2][3] == ("restart", "PT Mitra") and rows[2][4] == ("restart", "m@x")
        assert rows[3][3] == ("continue", "") and rows[3][4] == ("continue", "")

    def test_tanda_tetap_di_baris_yang_benar(self):
        tanda = [[True] * 4 for _ in L5.daftar_sub_api()]
        i_pay = [i for i, r in enumerate(L5.daftar_sub_api()) if r["sub_api"].endswith("Payment")][0]
        tanda[i_pay] = [False, False, True, True]
        doc = L5.build_lampiran_5b("SM-1", "u", tanda, sub_api_tampil=self.PILIH)
        for r in doc.tables[0].rows[L5.BARIS_DATA_MULAI:]:
            c = [x.text for x in L5._unique_cells(r)]
            exp = ([L5.TANDA_BELUM] * 2 + [L5.TANDA_SESUAI] * 2) if c[2].endswith("Payment") \
                else [L5.TANDA_SESUAI] * 4
            assert c[4:8] == exp

    def test_dokumen_valid_setelah_disimpan(self):
        doc = L6.build_lampiran_6b("PT", "u", sub_api_tampil={"API Virtual Account – Inquiry Status"})
        buf = io.BytesIO()
        doc.save(buf)
        d2 = Document(io.BytesIO(buf.getvalue()))
        assert [L5._unique_cells(r)[2].text for r in d2.tables[0].rows[L5.BARIS_DATA_MULAI:]] == [
            "API Virtual Account – Inquiry Status"]

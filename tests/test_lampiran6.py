"""
Lampiran 6A & 6B (Developer Site Pengguna Layanan): dibuat dari template;
hanya tanggal, penandatangan, Pengguna Layanan, User Devsite, Ѵ/X, dan
Layanan API yang digunakan yang diganti.
"""
import io
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from docx import Document

import lampiran5 as L5
import lampiran6 as L6

PEN = {"jabatan_1": "Jab A", "nama_1": "Nama A", "jabatan_2": "Jab B", "nama_2": "Nama B"}


def _uat_bytes(judul):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "UAT Script"
    ws.cell(row=2, column=2, value=judul)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _texts(doc):
    out = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for r in t.rows:
            out += [c.text for c in r.cells]
    return out


class TestNamaPengguna:
    def test_ambil_nama_pt_dari_judul(self):
        cases = {
            "Pengembangan dan Pengadaan Teknologi Informasi \nUser Acceptance Test Script\n"
            "<Penambahan Layanan Fund Transfer Host to Host PT Ezeelink Indonesia - SM-26-04-001-BSS-IS>":
                "PT Ezeelink Indonesia",
            "User Acceptance Test Script\n<Penambahan Layanan QRIS Merchant Aggregator PT Sender "
            "Integrasi Digital (Kirimo) (SM-26-07-004-BSS-IS)>":
                "PT Sender Integrasi Digital (Kirimo)",
            "User Acceptance Test Script\n<Penambahan Layanan Virtual Account & Fund Transfer Host "
            "to Host PT Solusi Pembayaran Elektronik (CRING)>":
                "PT Solusi Pembayaran Elektronik (CRING)",
            "User Acceptance Test Script\n<Penambahan Layanan QRIS Merchant Aggregator\n"
            "PT Cashlez Worldwide Indonesia (CashUp)>":
                "PT Cashlez Worldwide Indonesia (CashUp)",
        }
        for judul, exp in cases.items():
            assert L6.nama_pt_dari_judul(judul) == exp
            assert L6.read_nama_pengguna(_uat_bytes(judul)) == exp

    def test_tanpa_pt(self):
        assert L6.nama_pt_dari_judul("User Acceptance Test Script <Penambahan Layanan>") == ""


class TestLampiran6:
    def test_6a_hanya_tanggal_dan_ttd_yang_berubah(self):
        doc = L6.build_lampiran_6a(date(2026, 11, 9), PEN)
        tpl = _texts(Document(str(L6.TEMPLATE_6A)))
        out = _texts(doc)
        assert len(tpl) == len(out)
        txt = "\n".join(out)
        assert "Pada hari senin, tanggal 9 bulan november tahun 2026, telah selesai" in txt
        for v in PEN.values():
            assert v in txt
        assert "No." not in txt   # 6A memang tanpa nomor surat

    def test_6a_identik_template_bila_data_sama(self):
        pen = L5.default_penandatangan()
        doc = L6.build_lampiran_6a(date(2026, 7, 20), pen)
        tpl = _texts(Document(str(L6.TEMPLATE_6A)))
        assert _texts(doc) == tpl

    def test_6b_pengguna_user_tanda(self):
        rows = L6.daftar_sub_api()
        tanda = [[True] * 4 for _ in rows]
        tanda[0] = [True, False, True, True]
        doc = L6.build_lampiran_6b("PT Mitra", "a@mitra.id", tanda,
                                   layanan_tampil=["API Balance Inquiry", L5.LAYANAN_TSI])
        data = [[c.text for c in L5._unique_cells(r)] for r in doc.tables[0].rows[L5.BARIS_DATA_MULAI:]]
        assert [d[1] for d in data] == ["API Balance Inquiry", L5.LAYANAN_TSI]
        assert [d[0] for d in data] == ["1", "2"]
        for d in data:
            assert d[L6.KOL_PENGGUNA] == "PT Mitra" and d[L6.KOL_USER] == "a@mitra.id"
        assert data[0][5:9] == [L5.TANDA_SESUAI, L5.TANDA_BELUM, L5.TANDA_SESUAI, L5.TANDA_SESUAI]

    def test_6b_urutan_layanan_sama_dengan_5b(self):
        assert L6.daftar_layanan() == L5.daftar_layanan()
        assert [r["sub_api"] for r in L6.daftar_sub_api()] == [r["sub_api"] for r in L5.daftar_sub_api()]

    def test_6b_semua_layanan_jumlah_baris_sama_template(self):
        doc = L6.build_lampiran_6b("PT X", "x@x", layanan_tampil=L6.daftar_layanan())
        assert len(doc.tables[0].rows) == len(Document(str(L6.TEMPLATE_6B)).tables[0].rows)

    def test_convert_nama_otomatis_dan_qris_dilewati(self):
        src = _uat_bytes("User Acceptance Test Script\n<Penambahan Layanan PT Contoh Mitra - SM-1-BSS-IS>")
        res = L6.convert_lampiran_6_bytes(src, "FT_VA", PEN, "", "u@x", tanggal=date(2026, 1, 5))
        assert res["pengguna"] == "PT Contoh Mitra"
        assert res["a_name"].endswith("PT Contoh Mitra.docx")
        assert res["layanan"] == L6.daftar_layanan()      # tak ada section -> semua
        assert L6.convert_lampiran_6_bytes(src, "QRIS", PEN, "", "u") is None

    def test_convert_override_nama(self):
        src = _uat_bytes("User Acceptance Test Script\n<Penambahan Layanan PT Contoh Mitra>")
        res = L6.convert_lampiran_6_bytes(src, "FT_VA", PEN, "PT Nama Lain", "u@x")
        b = Document(io.BytesIO(res["b_bytes"]))
        assert "PT Nama Lain" in "\n".join(_texts(b))

"""
Lampiran 7A & 7B (Berita Acara Uji Fungsionalitas): dibuat dari template;
hanya tanggal, penandatangan, Pengguna Layanan, Ѵ/X, dan Sub API yang
digunakan yang diganti. 7B landscape & tabel presisi.
"""
import io
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn

import lampiran5 as L5
import lampiran7 as L7

PEN = {"jabatan_1": "Jab A", "nama_1": "Nama A", "jabatan_2": "Jab B", "nama_2": "Nama B"}


def _texts(doc):
    out = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for r in t.rows:
            out += [c.text for c in r.cells]
    return out


def _uat_bytes(judul="User Acceptance Test Script\n<Penambahan Layanan PT Contoh Mitra (CM)>"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "UAT Script"
    ws.cell(row=2, column=2, value=judul)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _data(doc):
    return [[c.text for c in L5._unique_cells(r)] for r in doc.tables[0].rows[L5.BARIS_DATA_MULAI:]]


class TestLampiran7A:
    def test_identik_template_bila_data_sama(self):
        doc = L7.build_lampiran_7a(date(2026, 7, 30), L5.default_penandatangan())
        assert _texts(doc) == _texts(Document(str(L7.TEMPLATE_7A)))

    def test_tanggal_dan_penandatangan(self):
        txt = "\n".join(_texts(L7.build_lampiran_7a(date(2026, 11, 9), PEN)))
        assert "Pada hari senin, tanggal 9 bulan november tahun 2026, telah selesai dilakukan uji" in txt
        for v in PEN.values():
            assert v in txt

    def test_tetap_portrait(self):
        doc = L7.build_lampiran_7a(date(2026, 1, 5), PEN)
        assert doc.sections[0].orientation == WD_ORIENT.PORTRAIT


class TestLampiran7B:
    def test_semua_sub_api_isi_sama_template(self):
        doc = L7.build_lampiran_7b("PT Silot Technology Indonesia")
        tpl = Document(str(L7.TEMPLATE_7B))
        assert [c.text for r in doc.tables[0].rows for c in r.cells] == \
               [c.text for r in tpl.tables[0].rows for c in r.cells]

    def test_set_sub_api_sama_dengan_5b(self):
        assert sorted(r["sub_api"] for r in L7.daftar_sub_api()) == \
               sorted(r["sub_api"] for r in L5.daftar_sub_api())

    def test_tanda_dicocokkan_per_nama_bukan_urutan(self):
        # Urutan 7B: TSI sebelum VA; list tanda mengikuti urutan 5B (TSI terakhir)
        tanda = [[True] * 4 for _ in L5.daftar_sub_api()]
        i_tsi = [i for i, r in enumerate(L5.daftar_sub_api()) if r["sub_api"] == L5.SUB_API_TSI][0]
        tanda[i_tsi] = [False, True, True, False]
        for row in _data(L7.build_lampiran_7b("PT X", tanda)):
            exp = [L5.TANDA_BELUM, L5.TANDA_SESUAI, L5.TANDA_SESUAI, L5.TANDA_BELUM] \
                if row[2] == L5.SUB_API_TSI else [L5.TANDA_SESUAI] * 4
            assert row[4:8] == exp

    def test_saring_sub_api_dan_nomor(self):
        pilih = {"API Balance Inquiry", "API Intrabank Transfer", L5.SUB_API_TSI,
                 "API Virtual Account – Payment"}
        data = _data(L7.build_lampiran_7b("PT X", sub_api_tampil=pilih))
        assert [d[2] for d in data] == ["API Balance Inquiry", "API Intrabank Transfer",
                                        L5.SUB_API_TSI, "API Virtual Account – Payment"]
        assert [d[0] for d in data] == ["1", "2", "3", "4"]
        assert all(d[L7.KOL_PENGGUNA] == "PT X" for d in data)

    def test_landscape_dan_lebar_tabel_pas(self):
        doc = L7.build_lampiran_7b("PT X")
        s = doc.sections[0]
        assert s.orientation == WD_ORIENT.LANDSCAPE
        area = round((s.page_width - s.left_margin - s.right_margin) / 635)
        tbl = doc.tables[0]._tbl
        assert int(tbl.tblPr.find(qn("w:tblW")).get(qn("w:w"))) == area
        assert sum(int(g.get(qn("w:w"))) for g in tbl.tblGrid.findall(qn("w:gridCol"))) == area


class TestConvert7:
    def test_nama_otomatis_dan_nama_file(self):
        res = L7.convert_lampiran_7_bytes(_uat_bytes(), "FT_VA", PEN, "", tanggal=date(2026, 1, 5))
        assert res["pengguna"] == "PT Contoh Mitra"
        assert res["a_name"] == "Lampiran 7A - Berita Acara Uji Fungsionalitas - PT Contoh Mitra.docx"
        assert res["b_name"].endswith("PT Contoh Mitra.docx")
        assert set(res["sub_api"]) == {r["sub_api"] for r in L7.daftar_sub_api()}   # tak ada section -> semua

    def test_qris_dilewati(self):
        assert L7.convert_lampiran_7_bytes(_uat_bytes(), "QRIS", PEN, "") is None

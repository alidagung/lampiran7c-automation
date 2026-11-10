"""
Lampiran 5B & 6B: halaman LANDSCAPE dan tabel presisi (pas di dalam margin,
lebar sel = grid, tinggi baris seragam, header berulang). 5A/6A tetap portrait.
"""
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn

import lampiran5 as L5
import lampiran6 as L6


def _cek_tabel(doc):
    s = doc.sections[0]
    assert s.orientation == WD_ORIENT.LANDSCAPE
    assert s.page_width > s.page_height
    area = round((s.page_width - s.left_margin - s.right_margin) / 635)
    tbl = doc.tables[0]._tbl
    grid = [int(g.get(qn("w:w"))) for g in tbl.tblGrid.findall(qn("w:gridCol"))]
    assert int(tbl.tblPr.find(qn("w:tblW")).get(qn("w:w"))) == area == sum(grid)
    assert tbl.tblPr.find(qn("w:tblInd")).get(qn("w:w")) == "0"
    assert tbl.tblPr.find(qn("w:jc")).get(qn("w:val")) == "center"
    tinggi = set()
    for ri, tr in enumerate(tbl.tr_lst):
        col = 0
        for tc in tr.tc_lst:
            sp = tc.tcPr.find(qn("w:gridSpan"))
            n = int(sp.get(qn("w:val"))) if sp is not None else 1
            assert int(tc.tcPr.find(qn("w:tcW")).get(qn("w:w"))) == sum(grid[col:col + n])
            col += n
        assert col == len(grid)
        h = tr.trPr.find(qn("w:trHeight"))
        tinggi.add((ri < L5.BARIS_DATA_MULAI, h.get(qn("w:val")), h.get(qn("w:hRule"))))
        if ri < L5.BARIS_DATA_MULAI:
            assert tr.trPr.find(qn("w:tblHeader")) is not None
    # satu ukuran untuk header, satu untuk data
    assert len({t for t in tinggi if t[0]}) == 1 and len({t for t in tinggi if not t[0]}) == 1


class TestLandscape:
    def test_5b_landscape_presisi(self):
        _cek_tabel(L5.build_lampiran_5b("SM-1", "u@x"))
        _cek_tabel(L5.build_lampiran_5b("SM-1", "u@x", layanan_tampil=["API Balance Inquiry"]))

    def test_6b_landscape_presisi(self):
        _cek_tabel(L6.build_lampiran_6b("PT X", "u@x"))
        _cek_tabel(L6.build_lampiran_6b("PT X", "u@x", layanan_tampil=[L5.LAYANAN_TSI]))

    def test_5a_6a_tetap_portrait(self):
        from datetime import date
        pen = L5.default_penandatangan()
        for doc in (L5.build_lampiran_5a("SM-1", date(2026, 1, 5), pen),
                    L6.build_lampiran_6a(date(2026, 1, 5), pen)):
            assert doc.sections[0].orientation == WD_ORIENT.PORTRAIT

    def test_isi_tidak_berubah_oleh_tata_letak(self):
        doc = L5.build_lampiran_5b("SM-1", "u@x")
        sub = [L5._unique_cells(r)[2].text for r in doc.tables[0].rows[L5.BARIS_DATA_MULAI:]]
        assert sub == [r["sub_api"] for r in L5.daftar_sub_api()]

"""
Lampiran 5A & 5B: dibuat dari template; hanya nomor surat, tanggal,
penandatangan, User DevSite, dan Ѵ/X yang diganti.
"""
import io
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from docx import Document

import lampiran5 as L


def _uat_bytes(nomor="SM-26-06-003-BSS-IS"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "UAT Script"
    ws.cell(row=6, column=2, value="Nomor Referensi")
    if nomor is not None:
        ws.cell(row=6, column=3, value=nomor)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _texts(doc):
    out = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for r in t.rows:
            out += [c.text for c in r.cells]
    return out


PEN = {"jabatan_1": "Jab A", "nama_1": "Nama A", "jabatan_2": "Jab B", "nama_2": "Nama B"}


class TestDataUat:
    def test_nomor_referensi_terbaca(self):
        assert L.read_nomor_referensi(_uat_bytes()) == "SM-26-06-003-BSS-IS"

    def test_nomor_referensi_kosong(self):
        assert L.read_nomor_referensi(_uat_bytes(None)) == ""

    def test_nama_hari_dihitung_dari_tanggal(self):
        # 30 Juli 2026 = Kamis (contoh template menulis "senin" -> salah)
        assert L.format_tanggal_pengujian(date(2026, 7, 30)) == "kamis, tanggal 30 bulan juli tahun 2026"
        assert L.format_tanggal_pengujian(date(2026, 11, 9)) == "senin, tanggal 9 bulan november tahun 2026"


class TestLampiran5:
    def _buat(self, **kw):
        res = L.convert_lampiran_5_bytes(_uat_bytes(), "FT_VA", PEN, "u@test.com",
                                         tanggal=date(2026, 11, 9), **kw)
        return res, Document(io.BytesIO(res["a_bytes"])), Document(io.BytesIO(res["b_bytes"]))

    def test_5a_nomor_tanggal_penandatangan(self):
        _, a, _ = self._buat()
        txt = "\n".join(_texts(a))
        assert "No. SM-26-06-003-BSS-IS" in txt
        assert "Pada hari senin, tanggal 9 bulan november tahun 2026, telah selesai" in txt
        for v in PEN.values():
            assert v in txt

    def test_5a_format_bold_nomor_dipertahankan(self):
        _, a, _ = self._buat()
        p = [p for p in a.paragraphs if p.text.startswith("No.")][0]
        nomor_run = [r for r in p.runs if "SM-26" in r.text][0]
        assert nomor_run.bold is True

    def test_5a_hanya_bagian_berubah_yang_diganti(self):
        _, a, _ = self._buat()
        tpl = _texts(Document(str(L.TEMPLATE_5A)))
        out = _texts(a)
        beda = [i for i, (x, y) in enumerate(zip(tpl, out)) if x != y]
        # nomor, kalimat tanggal, dan 4 sel penandatangan (tanpa duplikat sel)
        assert len(tpl) == len(out)
        assert len(beda) <= 6

    def test_5b_user_dan_tanda(self):
        n = len(L.daftar_sub_api())
        tanda = [[True] * 4 for _ in range(n)]
        tanda[1] = [True, True, False, True]
        _, _, b = self._buat(tanda=tanda)
        rows = b.tables[0].rows[L.BARIS_DATA_MULAI:]
        r1 = [c.text for c in L._unique_cells(rows[1])]
        assert r1[L.KOL_USER] == "u@test.com"
        assert r1[4:8] == [L.TANDA_SESUAI, L.TANDA_SESUAI, L.TANDA_BELUM, L.TANDA_SESUAI]
        assert [c.text for c in L._unique_cells(rows[0])][4:8] == [L.TANDA_SESUAI] * 4

    def test_5b_struktur_tabel_tetap(self):
        _, _, b = self._buat()
        tpl = Document(str(L.TEMPLATE_5B))
        assert b.element.xml.count("w:vMerge") == tpl.element.xml.count("w:vMerge")
        assert len(b.tables[0].rows) == len(tpl.tables[0].rows)

    def test_nomor_override_dan_nama_file(self):
        res, a, b = self._buat(nomor_surat="SM-99-01-001-BSS-IS")
        assert "No. SM-99-01-001-BSS-IS" in "\n".join(_texts(a))
        assert "No. SM-99-01-001-BSS-IS" in "\n".join(_texts(b))
        assert res["a_name"].endswith("SM-99-01-001-BSS-IS.docx")

    def test_qris_belum_didukung(self):
        assert L.convert_lampiran_5_bytes(_uat_bytes(), "QRIS", PEN, "u") is None

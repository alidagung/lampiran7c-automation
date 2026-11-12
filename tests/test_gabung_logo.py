"""
Penggabungan Lampiran 5A ... 7C menjadi 1 file, dan logo BSS di header UAT Result.
"""
import io
import os
import sys
import zipfile
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document
from docx.oxml.ns import qn

import gabung as G
import lampiran5 as L5
import lampiran6 as L6
import lampiran7 as L7
import main as m

BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
PEN = L5.default_penandatangan()


def _uat_mini():
    return {"Balance Services": [
        {"nomor_kasus_tes": "1.1", "langkah_tes": "Access Token Invalid", "hasil_aktual": "Berhasil",
         "nama_modul": "Balance Inquiry",
         "remarks": 'URL:\nhttps://x\n\nHeaders:\nA=b\n\nRequest Body:\n{"a":1}\n\nResponse:\n{"ok":1}'}]}


def _bagian():
    tgl = date(2026, 11, 9)
    b5a = L5.build_lampiran_5a("SM-1", tgl, PEN)
    b5b = L5.build_lampiran_5b("SM-1", "u@bank")
    b6a = L6.build_lampiran_6a(tgl, PEN)
    b6b = L6.build_lampiran_6b("PT X", "u@x")
    b7a = L7.build_lampiran_7a(tgl, PEN)
    b7b = L7.build_lampiran_7b("PT X")
    lamp = m.map_uat_to_lampiran(_uat_mini())
    b7c = m.build_lampiran_document(lamp)
    out = []
    for lab, d in (("5A", b5a), ("5B", b5b), ("6A", b6a), ("6B", b6b), ("7A", b7a), ("7B", b7b), ("7C", b7c)):
        buf = io.BytesIO()
        d.save(buf)
        out.append((lab, buf.getvalue()))
    return out


def _teks(doc):
    return [t.text for t in doc.element.body.iter(qn("w:t")) if t.text and t.text.strip()]


class TestGabung:
    def test_struktur_section_orientasi_header(self):
        bagian = _bagian()
        d = Document(io.BytesIO(G.gabung_bytes([b for _l, b in bagian], [l for l, _b in bagian])))
        assert len(d.sections) == 7
        orient = ["L" if s.page_width > s.page_height else "P" for s in d.sections]
        assert orient == ["P", "L", "P", "L", "P", "L", "L"]
        logo = [0 if s.header.is_linked_to_previous else len(list(s.header._element.iter(BLIP)))
                for s in d.sections]
        assert logo == [1, 1, 0, 0, 0, 0, 0]     # hanya 5A & 5B ber-logo (seperti template)

    def test_teks_tidak_hilang_dan_urutan_benar(self):
        bagian = _bagian()
        d = Document(io.BytesIO(G.gabung_bytes([b for _l, b in bagian], [l for l, _b in bagian])))
        semua = []
        for _l, b in bagian:
            semua += _teks(Document(io.BytesIO(b)))
        assert _teks(d) == semua

    def test_style_7c_tidak_tertimpa(self):
        # Normal 7C = 9pt; harus tetap 9pt walau master (5A) Normal = 11pt
        bagian = _bagian()
        d = Document(io.BytesIO(G.gabung_bytes([b for _l, b in bagian], [l for l, _b in bagian])))
        styles = {s.get(qn("w:styleId")): s for s in d.styles.element.findall(qn("w:style"))}
        normal_7c = styles["L7C_Normal"]
        assert normal_7c.find(qn("w:rPr")).find(qn("w:sz")).get(qn("w:val")) == "18"
        ids = {e.get(qn("w:val")) for e in d.element.body.iter(qn("w:pStyle"))}
        assert ids <= set(styles)

    def test_paket_valid(self):
        bagian = _bagian()
        gb = G.gabung_bytes([b for _l, b in bagian], [l for l, _b in bagian])
        z = zipfile.ZipFile(io.BytesIO(gb))
        assert len(z.namelist()) == len(set(z.namelist()))
        names = [s.find(qn("w:name")).get(qn("w:val")) for s in
                 Document(io.BytesIO(gb)).styles.element.findall(qn("w:style"))]
        assert len(names) == len(set(names))

    def test_satu_dokumen_saja(self):
        lab, b = _bagian()[-1]
        d = Document(io.BytesIO(G.gabung_bytes([b], [lab])))
        assert len(d.sections) == 1


class TestLogoUatResult:
    def test_logo_di_header_kanan(self):
        doc = m.build_uat_result_document(_uat_mini())
        h = doc.sections[0].header
        assert len(list(h._element.iter(BLIP))) == 1
        assert h.paragraphs[0].alignment == 2   # RIGHT
        ext = h._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent")
        assert abs(int(ext.get("cx")) / 360000 - m.LOGO_BSS_LEBAR_CM) < 0.01

    def test_tanpa_file_logo_tetap_jalan(self, tmp_path):
        doc = m.Document()
        assert m._pasang_logo_header(doc, logo=tmp_path / "tidak_ada.png") is False

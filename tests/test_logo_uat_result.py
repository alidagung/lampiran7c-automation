"""Logo BSS di header UAT Result."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import main as m

BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"


def _uat_mini():
    return {"Balance Services": [
        {"nomor_kasus_tes": "1.1", "langkah_tes": "Access Token Invalid", "hasil_aktual": "Berhasil",
         "nama_modul": "Balance Inquiry",
         "remarks": 'URL:\nhttps://x\n\nHeaders:\nA=b\n\nRequest Body:\n{"a":1}\n\nResponse:\n{"ok":1}'}]}


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

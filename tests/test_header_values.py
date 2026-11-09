"""
Prinsip: VALUE header tidak boleh diubah. Hanya sintaks pembungkus (kutip
pembungkus item, koma pemisah item) yang boleh dibuang.
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import _format_headers

CONTOH = (
    '[\n'
    '  "content-type=application/json",\n'
    '  "X-PARTNER-ID=x8ir2mpkj477amer19biu4kn2l8cwlcw",\n'
    '  "Authorization=Bearer test",\n'
    '  "Cookie=Path=/"\n'
    ']'
)


class TestNilaiHeaderTidakBerubah:
    def test_format_sesuai_contoh_dari_semua_bentuk_input(self):
        plain = ("content-type=application/json\nX-PARTNER-ID=x8ir2mpkj477amer19biu4kn2l8cwlcw\n"
                 "Authorization=Bearer test\nCookie=Path=/")
        colon = ("content-type: application/json\nX-PARTNER-ID: x8ir2mpkj477amer19biu4kn2l8cwlcw\n"
                 "Authorization: Bearer test\nCookie: Path=/")
        quoted = ('"content-type=application/json",\n"X-PARTNER-ID=x8ir2mpkj477amer19biu4kn2l8cwlcw",\n'
                  '"Authorization=Bearer test",\n"Cookie=Path=/"')
        for raw in (plain, colon, quoted):
            assert _format_headers(raw) == CONTOH

    def test_huruf_kecil_dan_tanda_sama_dengan_dalam_value_dipertahankan(self):
        out = _format_headers("content-type=Application/JSON\nCookie=Path=/")
        assert '"content-type=Application/JSON"' in out
        assert '"Cookie=Path=/"' in out

    def test_koma_di_akhir_value_tanpa_kutip_dipertahankan(self):
        assert '"X-A=abc,"' in _format_headers("X-A=abc,")

    def test_kutip_di_dalam_value_dipertahankan(self):
        assert '"X-B=it"s"' in _format_headers('X-B: it"s')

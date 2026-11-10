"""
Lampiran 7A & 7B ASPI (Berita Acara Uji Fungsionalitas)
=======================================================
Sama seperti Lampiran 5/6, dokumen dibuat dari TEMPLATE resmi (templates/) dan
hanya bagian yang berubah yang diganti:

  Lampiran 7A : kalimat hari/tanggal pengujian, penandatangan
                (tidak ada nomor surat)
  Lampiran 7B : Pengguna Layanan (nama PT mitra), isian Ѵ / X, dan Sub API
                yang digunakan (hasil validasi UAT Script)

Catatan: urutan baris di template 7B dipertahankan apa adanya (API Transaction
Status Inquiry berada SEBELUM API Virtual Account). Karena itu isian Ѵ/X
dicocokkan berdasarkan NAMA Sub API, bukan nomor baris.
"""

from datetime import date

from docx import Document

from lampiran5 import (
    TEMPLATE_DIR, PRODUK_DIDUKUNG, TANDA_SESUAI, TANDA_BELUM, BARIS_DATA_MULAI,
    format_tanggal_pengujian, _replace_regex, _set_cell_text, _unique_cells,
    _cari_paragraf, _to_bytes, nama_file_aman, _set_landscape, _rapikan_tabel,
    saring_sub_api, pilih_sub_api, resolve_sub_api,
    daftar_sub_api as _daftar_sub_api_5,
)
from lampiran6 import read_nama_pengguna

TEMPLATE_7A = TEMPLATE_DIR / "Lampiran 7A.docx"
TEMPLATE_7B = TEMPLATE_DIR / "Lampiran 7B.docx"

# Kolom matriks 7B (indeks sel unik pada baris data)
KOL_PENGGUNA = 3
KOL_TANDA = (4, 5, 6, 7)

# Bobot lebar kolom grid 7B (proporsional thd lebar area tulis landscape):
# No | Layanan API | Sub API | Pengguna Layanan | Pos Req | Pos Resp | Neg Req | Neg Resp
BOBOT_KOLOM_7B = (520, 2900, 3400, 3150, 1100, 1100, 1100, 1100)


# ============================================================
# DATA TEMPLATE
# ============================================================

def daftar_sub_api():
    """Baris matriks 7B sesuai urutan template: list dict {layanan_api, sub_api}."""
    t = Document(str(TEMPLATE_7B)).tables[0]
    rows = []
    for row in t.rows[BARIS_DATA_MULAI:]:
        cells = _unique_cells(row)
        rows.append({"layanan_api": cells[1].text.strip(), "sub_api": cells[2].text.strip()})
    return rows


def daftar_layanan():
    out = []
    for r in daftar_sub_api():
        if r["layanan_api"] not in out:
            out.append(r["layanan_api"])
    return out


def _tanda_per_nama(tanda):
    """
    Ubah isian Ѵ/X menjadi dict {sub_api: [4 bool]}.
    `tanda` boleh berupa dict, atau list yang sejajar dengan urutan Sub API
    Lampiran 5B (urutan yang dipakai tabel validasi di aplikasi).
    """
    if tanda is None:
        return {}
    if isinstance(tanda, dict):
        return dict(tanda)
    return {r["sub_api"]: list(f) for r, f in zip(_daftar_sub_api_5(), tanda)}


# ============================================================
# PEMBUAT DOKUMEN
# ============================================================

def build_lampiran_7a(tanggal, penandatangan):
    doc = Document(str(TEMPLATE_7A))
    p_tgl = _cari_paragraf(doc, r"Pada hari")
    if p_tgl is not None:
        _replace_regex(
            p_tgl,
            r"Pada hari\s+(\S+,\s*tanggal\s+\d+\s+bulan\s+\S+\s+tahun\s+\d{4})",
            format_tanggal_pengujian(tanggal),
            group=1,
        )
    t = doc.tables[0]
    _set_cell_text(t.rows[0].cells[0], penandatangan.get("jabatan_1", ""))
    _set_cell_text(t.rows[0].cells[1], penandatangan.get("jabatan_2", ""))
    _set_cell_text(t.rows[2].cells[0], penandatangan.get("nama_1", ""))
    _set_cell_text(t.rows[2].cells[1], penandatangan.get("nama_2", ""))
    return doc


def build_lampiran_7b(pengguna_layanan, tanda=None, layanan_tampil=None, sub_api_tampil=None):
    """
    Args:
        pengguna_layanan: nama PT mitra (diisi ke semua baris)
        tanda           : isian Ѵ/X (dict per nama Sub API, atau list sejajar
                          urutan Lampiran 5B); None = semua Ѵ
        layanan_tampil  : Layanan API yang ditampilkan (semua Sub API-nya)
        sub_api_tampil  : Sub API yang ditampilkan; keduanya None = semua
    """
    doc = Document(str(TEMPLATE_7B))
    t = doc.tables[0]
    per_nama = _tanda_per_nama(tanda)
    for row in t.rows[BARIS_DATA_MULAI:]:
        cells = _unique_cells(row)
        _set_cell_text(cells[KOL_PENGGUNA], pengguna_layanan)
        flags = per_nama.get(cells[2].text.strip(), [True] * 4)
        for col, ok in zip(KOL_TANDA, flags):
            _set_cell_text(cells[col], TANDA_SESUAI if ok else TANDA_BELUM)

    rows = daftar_sub_api()
    saring_sub_api(t, rows, pilih_sub_api(rows, layanan_tampil, sub_api_tampil))

    # Halaman landscape + tabel presisi
    _set_landscape(doc)
    _rapikan_tabel(doc, t, BOBOT_KOLOM_7B)
    return doc


def convert_lampiran_7_bytes(source, product_key, penandatangan, pengguna_layanan,
                             tanda=None, tanggal=None, layanan_tampil=None,
                             sub_api_tampil=None):
    """
    Buat Lampiran 7A & 7B (bytes).

    Args:
        source          : path/bytes UAT Script (untuk validasi Sub API & nama PT)
        product_key     : "FT_VA"/"QRIS"; selain PRODUK_DIDUKUNG -> None
        penandatangan   : dict jabatan/nama (sama dgn Lampiran 5A)
        pengguna_layanan: nama PT mitra ("" -> dibaca dari judul UAT Script)
        tanda           : isian Ѵ/X (lihat build_lampiran_7b)
        tanggal         : date; default hari ini
        layanan_tampil  : Layanan API 7B (semua Sub API-nya)
        sub_api_tampil  : Sub API 7B; keduanya None = hasil validasi UAT Script

    Returns:
        dict {pengguna, tanggal, layanan, sub_api, a_bytes, a_name, b_bytes, b_name}
        atau None bila produk belum didukung.
    """
    if product_key not in PRODUK_DIDUKUNG:
        return None
    tanggal = tanggal or date.today()
    pengguna = (pengguna_layanan or "").strip() or read_nama_pengguna(source)
    sub_tampil = resolve_sub_api(source, layanan_tampil, sub_api_tampil)

    doc_a = build_lampiran_7a(tanggal, penandatangan)
    doc_b = build_lampiran_7b(pengguna, tanda, sub_api_tampil=sub_tampil)
    rows = daftar_sub_api()
    layanan = []
    for r in rows:
        if r["sub_api"] in sub_tampil and r["layanan_api"] not in layanan:
            layanan.append(r["layanan_api"])
    akhiran = f" - {nama_file_aman(pengguna)}" if pengguna else ""
    return {
        "pengguna": pengguna,
        "tanggal": tanggal,
        "layanan": layanan,
        "sub_api": [r["sub_api"] for r in rows if r["sub_api"] in sub_tampil],
        "a_bytes": _to_bytes(doc_a),
        "a_name": f"Lampiran 7A - Berita Acara Uji Fungsionalitas{akhiran}.docx",
        "b_bytes": _to_bytes(doc_b),
        "b_name": f"Lampiran 7B - Matriks Asesmen Uji Fungsionalitas{akhiran}.docx",
    }

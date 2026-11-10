"""
Lampiran 6A & 6B ASPI (Informasi Uji Developer Site Pengguna Layanan)
=====================================================================
Sama seperti Lampiran 5A/5B, dokumen dibuat dari TEMPLATE resmi (templates/)
dan hanya bagian yang berubah yang diganti:

  Lampiran 6A : kalimat hari/tanggal pengujian, penandatangan
                (tidak ada nomor surat)
  Lampiran 6B : Pengguna Layanan (nama PT mitra), User pada Devsite (email
                mitra), isian Ѵ / X, dan Layanan API yang digunakan

Sumber data:
  - Pengguna Layanan : nama PT pada judul UAT Script (bisa diedit di aplikasi)
  - User Devsite     : diketik di aplikasi (email mitra, beda per mitra)
  - Tanggal          : tanggal saat diproses
  - Penandatangan    : sama dengan pengaturan Lampiran 5A
  - Layanan API      : sama dengan aturan Lampiran 5B (yang digunakan mitra)
"""

import copy
import io
import re
from datetime import date

from docx import Document
from docx.table import _Cell
from openpyxl import load_workbook

from lampiran5 import (
    TEMPLATE_DIR, PRODUK_DIDUKUNG, TANDA_SESUAI, TANDA_BELUM,
    LAYANAN_TSI, SUB_API_TSI, BARIS_CONTOH_TUNGGAL, BARIS_DATA_MULAI,
    format_tanggal_pengujian, layanan_digunakan, daftar_layanan as _daftar_layanan_5,
    _replace_regex, _set_cell_text, _unique_cells, _cari_paragraf, _to_bytes,
    nama_file_aman,
)

TEMPLATE_6A = TEMPLATE_DIR / "Lampiran 6A.docx"
TEMPLATE_6B = TEMPLATE_DIR / "Lampiran 6B.docx"

# Kolom matriks 6B (indeks sel unik pada baris data)
KOL_PENGGUNA = 3
KOL_USER = 4
KOL_TANDA = (5, 6, 7, 8)


# ============================================================
# DATA DARI UAT SCRIPT
# ============================================================

def nama_pt_dari_judul(judul):
    """
    Ambil nama PT mitra dari judul UAT Script, contoh:
      "<Penambahan Layanan Fund Transfer Host to Host PT Ezeelink Indonesia - SM-26-04-001-BSS-IS>"
        -> "PT Ezeelink Indonesia"
      "<... PT Sender Integrasi Digital (Kirimo) (SM-26-07-004-BSS-IS)>"
        -> "PT Sender Integrasi Digital"
    Hanya NAMA PT yang diambil: nomor SM dan semua keterangan dalam kurung
    (singkatan / nama brand) dibuang. Mengembalikan "" bila tidak ada "PT".
    """
    teks = re.sub(r"\s+", " ", str(judul or "").replace("<", " ").replace(">", " ")).strip()
    m = re.search(r"\bPT\b.*$", teks)
    if not m:
        return ""
    nama = m.group(0).strip()
    # Buang nomor SM di akhir: " - SM-..." atau " (SM-...)"
    nama = re.sub(r"\s*[-–]\s*SM-[\w-]+.*$", "", nama)
    # Buang semua keterangan dalam kurung (singkatan/nama brand/nomor SM),
    # mis. "(CRING)", "(Pipo Grup)", "(FT)", "(Kirimo)", "(SM-26-...)".
    nama = re.sub(r"\s*\([^()]*\)", "", nama)
    nama = re.sub(r"\s+", " ", nama)
    return nama.strip(" -–")


def read_nama_pengguna(source):
    """Nama PT mitra dari judul (baris paling atas) sheet UAT Script."""
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    elif hasattr(source, "seek"):
        source.seek(0)
    wb = load_workbook(source, read_only=True, data_only=True)
    try:
        ws = wb["UAT Script"] if "UAT Script" in wb.sheetnames else wb.active
        for row in ws.iter_rows(min_row=1, max_row=12, values_only=True):
            for v in row or ():
                if v and "user acceptance test" in str(v).lower():
                    nama = nama_pt_dari_judul(v)
                    if nama:
                        return nama
        return ""
    finally:
        wb.close()


# ============================================================
# TEMPLATE 6B (dinormalisasi seperti 5B)
# ============================================================

def _template_6b():
    """
    Muat template 6B lalu keluarkan baris "API Transaction Status Inquiry" dari
    grup Virtual Account menjadi baris tersendiri di akhir tabel (konsisten
    dengan Lampiran 5B).
    """
    doc = Document(str(TEMPLATE_6B))
    table = doc.tables[0]
    trs = list(table._tbl.tr_lst)[BARIS_DATA_MULAI:]

    def teks(tc):
        return _Cell(tc, table).text.strip()

    tr_tsi = next((tr for tr in trs if teks(tr.tc_lst[2]) == SUB_API_TSI), None)
    tr_contoh = next((tr for tr in trs if teks(tr.tc_lst[1]) == BARIS_CONTOH_TUNGGAL), None)
    if tr_tsi is None or tr_contoh is None:
        return doc

    baru = copy.deepcopy(tr_contoh)
    tr_tsi.getparent().remove(tr_tsi)
    table._tbl.append(baru)
    _set_cell_text(_Cell(baru.tc_lst[1], table), LAYANAN_TSI)
    _set_cell_text(_Cell(baru.tc_lst[2], table), SUB_API_TSI)
    return doc


def daftar_sub_api():
    """Baris matriks 6B: list of dict {layanan_api, sub_api}."""
    t = _template_6b().tables[0]
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


# ============================================================
# PEMBUAT DOKUMEN
# ============================================================

def build_lampiran_6a(tanggal, penandatangan):
    doc = Document(str(TEMPLATE_6A))
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


def build_lampiran_6b(pengguna_layanan, user_devsite, tanda=None, layanan_tampil=None):
    """
    Args:
        pengguna_layanan: nama PT mitra (diisi ke semua baris)
        user_devsite    : email user DevSite mitra (diisi ke semua baris)
        tanda           : list per baris sub API berisi 4 bool (True=Ѵ, False=X)
        layanan_tampil  : Layanan API yang ditampilkan (None = semua)
    """
    doc = _template_6b()
    t = doc.tables[0]
    for i, row in enumerate(t.rows[BARIS_DATA_MULAI:]):
        cells = _unique_cells(row)
        _set_cell_text(cells[KOL_PENGGUNA], pengguna_layanan)
        _set_cell_text(cells[KOL_USER], user_devsite)
        flags = tanda[i] if tanda is not None and i < len(tanda) else [True] * 4
        for col, ok in zip(KOL_TANDA, flags):
            _set_cell_text(cells[col], TANDA_SESUAI if ok else TANDA_BELUM)

    tampil = set(layanan_tampil) if layanan_tampil is not None else set(daftar_layanan())
    rows = daftar_sub_api()
    trs = list(t._tbl.tr_lst)[BARIS_DATA_MULAI:]
    for info, tr in zip(rows, trs):
        if info["layanan_api"] not in tampil:
            tr.getparent().remove(tr)
    nomor, sebelumnya = 0, None
    for info, tr in zip(rows, trs):
        if info["layanan_api"] not in tampil:
            continue
        if info["layanan_api"] != sebelumnya:
            nomor += 1
            sebelumnya = info["layanan_api"]
            _set_cell_text(_Cell(tr.tc_lst[0], t), str(nomor))
    return doc


def convert_lampiran_6_bytes(source, product_key, penandatangan, pengguna_layanan,
                             user_devsite, tanda=None, tanggal=None,
                             layanan_tampil=None):
    """
    Buat Lampiran 6A & 6B (bytes).

    Args:
        source          : path/bytes UAT Script (untuk Layanan API yang digunakan)
        product_key     : "FT_VA"/"QRIS"; selain PRODUK_DIDUKUNG -> None
        penandatangan   : dict jabatan/nama (sama dgn Lampiran 5A)
        pengguna_layanan: nama PT mitra ("" -> dibaca dari judul UAT Script)
        user_devsite    : email user DevSite mitra
        tanda           : isian Ѵ/X per baris (urutan daftar_sub_api())
        tanggal         : date; default hari ini
        layanan_tampil  : Layanan API 6B; default = yang digunakan menurut UAT

    Returns:
        dict {pengguna, tanggal, layanan, a_bytes, a_name, b_bytes, b_name}
        atau None bila produk belum didukung.
    """
    if product_key not in PRODUK_DIDUKUNG:
        return None
    tanggal = tanggal or date.today()
    pengguna = (pengguna_layanan or "").strip() or read_nama_pengguna(source)

    if layanan_tampil is None:
        import main
        src = io.BytesIO(source) if isinstance(source, (bytes, bytearray)) else source
        if hasattr(src, "seek"):
            src.seek(0)
        layanan_tampil = layanan_digunakan(main.read_uat_script(src))
    if not layanan_tampil:
        layanan_tampil = daftar_layanan()

    doc_a = build_lampiran_6a(tanggal, penandatangan)
    doc_b = build_lampiran_6b(pengguna, user_devsite, tanda, layanan_tampil)
    akhiran = f" - {nama_file_aman(pengguna)}" if pengguna else ""
    return {
        "pengguna": pengguna,
        "tanggal": tanggal,
        "layanan": list(layanan_tampil),
        "a_bytes": _to_bytes(doc_a),
        "a_name": f"Lampiran 6A - Informasi Uji Developer Site Pengguna Layanan{akhiran}.docx",
        "b_bytes": _to_bytes(doc_b),
        "b_name": f"Lampiran 6B - Matriks Asesmen Pengguna Layanan{akhiran}.docx",
    }

"""
Lampiran 5A & 5B ASPI (Informasi Uji Developer Site Penyedia Layanan)
=====================================================================
Dokumen dibuat dari TEMPLATE resmi (folder templates/) supaya format, font,
dan tabel persis sama dengan contoh. Hanya bagian yang berubah yang diganti:

  Lampiran 5A : nomor surat, kalimat hari/tanggal pengujian, penandatangan
  Lampiran 5B : nomor surat, User pada DevSite, isian Ѵ / X per sub API

Sumber data:
  - Nomor surat  : "Nomor Referensi" pada UAT Script (mis. SM-26-06-003-BSS-IS)
  - Tanggal      : tanggal saat dokumen dibuat (saat UAT diunggah & diproses)
  - Penandatangan, User DevSite, Ѵ/X : diisi di aplikasi (default = isi template)

Saat ini daftar Sub API mengikuti template (produk Fund Transfer / VA). Pada 5B
hanya Layanan API yang DIGUNAKAN mitra (menurut UAT Script) yang ditampilkan.
"""

import io
import re
from datetime import date
from pathlib import Path

from docx import Document
from docx.table import _Cell
from openpyxl import load_workbook

TEMPLATE_DIR = Path(__file__).resolve().parent / "templates"
TEMPLATE_5A = TEMPLATE_DIR / "Lampiran 5A.docx"
TEMPLATE_5B = TEMPLATE_DIR / "Lampiran 5B.docx"

# Tanda pada matriks 5B (karakter persis seperti template)
TANDA_SESUAI = "Ѵ"
TANDA_BELUM = "X"

# Produk yang sudah didukung (daftar Sub API di template = produk VA/FT)
PRODUK_DIDUKUNG = ("FT_VA",)

NAMA_HARI = ["senin", "selasa", "rabu", "kamis", "jumat", "sabtu", "minggu"]
NAMA_BULAN = ["januari", "februari", "maret", "april", "mei", "juni", "juli",
              "agustus", "september", "oktober", "november", "desember"]

# Layanan API di matriks 5B -> section terkait di UAT Script.
# Layanan dianggap DIGUNAKAN bila section-nya memiliki minimal satu case yang
# statusnya BUKAN "Tidak dites" (Tidak dites = mitra tidak menggunakan layanan).
LAYANAN_KE_SECTION = {
    "API Balance Inquiry": ("Balance Services",),
    "API Intrabank Transfer": ("Intrabank Transfer",),
    "API Interbank Transfer": ("Interbank Transfer",),
    "API RTGS Transfer": ("RTGS Transfer",),
    "API SKNBI Transfer": ("SKNBI Transfer",),
    "API Virtual Account": ("Transfer VA",),
}

# Kolom matriks 5B (indeks sel pada baris data)
KOL_USER = 3
KOL_TANDA = (4, 5, 6, 7)   # Positif Request, Positif Response, Negatif Request, Negatif Response
BARIS_DATA_MULAI = 2       # baris 0-1 = header bertingkat


# ============================================================
# DATA DARI UAT SCRIPT
# ============================================================

def read_nomor_referensi(source):
    """
    Ambil "Nomor Referensi" (nomor project, mis. SM-26-06-003-BSS-IS) dari blok
    metadata di atas tabel UAT Script. Mengembalikan "" bila tidak ditemukan.

    Args:
        source: path file ATAU bytes/file-like berisi .xlsx
    """
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    elif hasattr(source, "seek"):
        source.seek(0)
    wb = load_workbook(source, read_only=True, data_only=True)
    try:
        ws = wb["UAT Script"] if "UAT Script" in wb.sheetnames else wb.active
        for row in ws.iter_rows(min_row=1, max_row=40, values_only=True):
            if not row or len(row) < 3:
                continue
            label = str(row[1]).strip().lower() if row[1] is not None else ""
            if label == "nomor referensi" and row[2] is not None and str(row[2]).strip():
                return str(row[2]).strip()
        return ""
    finally:
        wb.close()


def format_tanggal_pengujian(tgl):
    """
    Kalimat tanggal sesuai gaya template, mis.
    date(2026, 7, 30) -> "kamis, tanggal 30 bulan juli tahun 2026"
    """
    return (f"{NAMA_HARI[tgl.weekday()]}, tanggal {tgl.day} "
            f"bulan {NAMA_BULAN[tgl.month - 1]} tahun {tgl.year}")


def layanan_digunakan(uat_data):
    """
    Daftar Layanan API (urut sesuai template 5B) yang digunakan mitra menurut
    UAT Script.

    Args:
        uat_data: dict hasil main.read_uat_script()
    """
    hasil = []
    for layanan in daftar_layanan():
        sections = LAYANAN_KE_SECTION.get(layanan, ())
        dipakai = any(
            str(r.get("hasil_aktual", "")).strip().lower() not in ("", "tidak dites")
            for s in sections for r in uat_data.get(s, [])
        )
        if dipakai:
            hasil.append(layanan)
    return hasil


# ============================================================
# UTILITAS PENGGANTIAN TEKS (mempertahankan format run)
# ============================================================

def _replace_span(paragraph, start, end, new_text):
    """
    Ganti teks paragraf pada posisi karakter [start, end) dengan new_text.
    Teks baru ditaruh di run pertama yang tersentuh (format run itu dipakai);
    run lain dalam rentang dikosongkan. Run di luar rentang tidak disentuh.
    """
    pos = 0
    placed = False
    for run in paragraph.runs:
        r_start, r_end = pos, pos + len(run.text)
        pos = r_end
        if r_end <= start or r_start >= end:
            continue
        before = run.text[:max(0, start - r_start)]
        after = run.text[min(len(run.text), end - r_start):] if r_end > end else ""
        if not placed:
            run.text = before + new_text + after
            placed = True
        else:
            run.text = before + after


def _replace_regex(paragraph, pattern, new_text, group=0):
    """Ganti kecocokan pertama `pattern` (pada group tertentu) di paragraf."""
    m = re.search(pattern, paragraph.text)
    if not m:
        return False
    _replace_span(paragraph, m.start(group), m.end(group), new_text)
    return True


def _set_cell_text(cell, text):
    """Isi sel dengan teks, memakai format run pertama yang ada."""
    para = cell.paragraphs[0]
    runs = para.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r.text = ""
    else:
        para.add_run(text)
    for extra in cell.paragraphs[1:]:
        for r in extra.runs:
            r.text = ""


def _unique_cells(row):
    """Sel unik pada satu baris (sel gabungan horizontal tidak diulang)."""
    seen, out = set(), []
    for c in row.cells:
        if id(c._tc) in seen:
            continue
        seen.add(id(c._tc))
        out.append(c)
    return out


def _cari_paragraf(doc, pola):
    for p in doc.paragraphs:
        if re.search(pola, p.text):
            return p
    return None


# ============================================================
# NILAI DEFAULT (dibaca dari template)
# ============================================================

def default_penandatangan():
    """Jabatan & nama penandatangan sesuai template 5A (untuk nilai awal form)."""
    t = Document(str(TEMPLATE_5A)).tables[0]
    return {
        "jabatan_1": t.rows[0].cells[0].text.strip(),
        "jabatan_2": t.rows[0].cells[1].text.strip(),
        "nama_1": t.rows[2].cells[0].text.strip(),
        "nama_2": t.rows[2].cells[1].text.strip(),
    }


def daftar_sub_api():
    """
    Daftar baris matriks 5B sesuai template:
    list of dict {no, layanan_api, sub_api, user}.
    """
    t = Document(str(TEMPLATE_5B)).tables[0]
    rows = []
    for row in t.rows[BARIS_DATA_MULAI:]:
        cells = _unique_cells(row)
        rows.append({
            "no": cells[0].text.strip(),
            "layanan_api": cells[1].text.strip(),
            "sub_api": cells[2].text.strip(),
            "user": cells[KOL_USER].text.strip(),
        })
    return rows


def daftar_layanan():
    """Layanan API unik pada template 5B (urutan sesuai template)."""
    out = []
    for r in daftar_sub_api():
        if r["layanan_api"] not in out:
            out.append(r["layanan_api"])
    return out


def default_user_devsite():
    rows = daftar_sub_api()
    return rows[0]["user"] if rows else ""


# ============================================================
# PEMBUAT DOKUMEN
# ============================================================

def build_lampiran_5a(nomor_surat, tanggal, penandatangan):
    """
    Args:
        nomor_surat  : str, mis. "SM-26-06-003-BSS-IS"
        tanggal      : datetime.date tanggal pengujian
        penandatangan: dict jabatan_1, nama_1, jabatan_2, nama_2
    """
    doc = Document(str(TEMPLATE_5A))

    p_no = _cari_paragraf(doc, r"^\s*No\.")
    if p_no is not None:
        _replace_regex(p_no, r"^\s*No\.\s*(\S.*?)\s*$", nomor_surat, group=1)

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


def build_lampiran_5b(nomor_surat, user_devsite, tanda=None, layanan_tampil=None):
    """
    Args:
        nomor_surat : str
        user_devsite: str, email user DevSite (diisi ke semua baris)
        tanda       : list (per baris sub API) berisi 4 bool
                      [positif_req, positif_resp, negatif_req, negatif_resp];
                      True -> Ѵ, False -> X. None = semua Ѵ.
        layanan_tampil: daftar Layanan API yang ditampilkan (lainnya dihapus
                      beserta semua Sub API-nya). None = semua.
    """
    doc = Document(str(TEMPLATE_5B))

    p_no = _cari_paragraf(doc, r"^\s*No\.")
    if p_no is not None:
        _replace_regex(p_no, r"^\s*No\.\s*(\S.*?)\s*$", nomor_surat, group=1)

    t = doc.tables[0]
    data_rows = t.rows[BARIS_DATA_MULAI:]
    for i, row in enumerate(data_rows):
        cells = _unique_cells(row)
        _set_cell_text(cells[KOL_USER], user_devsite)
        flags = tanda[i] if tanda is not None and i < len(tanda) else [True] * 4
        for col, ok in zip(KOL_TANDA, flags):
            _set_cell_text(cells[col], TANDA_SESUAI if ok else TANDA_BELUM)

    if layanan_tampil is not None:
        _saring_layanan(t, set(layanan_tampil))
    return doc


def _saring_layanan(table, tampil):
    """
    Hapus baris Layanan API yang tidak ditampilkan (satu grup utuh beserta
    semua Sub API-nya, sehingga sel gabungan No/Layanan tetap utuh), lalu
    nomori ulang kolom No secara berurutan (1, 2, 3, ...).
    """
    rows = daftar_sub_api()
    trs = list(table._tbl.tr_lst)[BARIS_DATA_MULAI:]
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
            _set_cell_text(_Cell(tr.tc_lst[0], table), str(nomor))


def _to_bytes(doc):
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def nama_file_aman(teks):
    return re.sub(r'[\\/:*?"<>|]+', "-", teks).strip()


def convert_lampiran_5_bytes(source, product_key, penandatangan, user_devsite,
                             tanda=None, tanggal=None, nomor_surat=None,
                             layanan_tampil=None):
    """
    Buat Lampiran 5A & 5B (bytes) dari file UAT Script.

    Args:
        source       : path/bytes file UAT Script (untuk Nomor Referensi)
        product_key  : "FT_VA" / "QRIS"; selain PRODUK_DIDUKUNG -> None
        penandatangan: dict jabatan/nama (lihat build_lampiran_5a)
        user_devsite : email user DevSite
        tanda        : isian Ѵ/X per baris (lihat build_lampiran_5b)
        tanggal      : date; default = hari ini (saat diproses)
        nomor_surat  : override manual; default = Nomor Referensi di UAT
        layanan_tampil: Layanan API yang ditampilkan di 5B; default = yang
                       digunakan menurut UAT Script (layanan_digunakan)

    Returns:
        dict {nomor_surat, tanggal, layanan, a_bytes, a_name, b_bytes, b_name}
        atau None bila produk belum didukung.
    """
    if product_key not in PRODUK_DIDUKUNG:
        return None
    tanggal = tanggal or date.today()
    nomor = (nomor_surat or "").strip() or read_nomor_referensi(source)

    if layanan_tampil is None:
        import main  # impor lokal (main tidak bergantung pada modul ini)
        if isinstance(source, (bytes, bytearray)):
            uat_src = io.BytesIO(source)
        else:
            uat_src = source
            if hasattr(uat_src, "seek"):
                uat_src.seek(0)
        layanan_tampil = layanan_digunakan(main.read_uat_script(uat_src))

    doc_a = build_lampiran_5a(nomor, tanggal, penandatangan)
    doc_b = build_lampiran_5b(nomor, user_devsite, tanda, layanan_tampil)
    akhiran = f" - {nama_file_aman(nomor)}" if nomor else ""
    return {
        "nomor_surat": nomor,
        "tanggal": tanggal,
        "layanan": list(layanan_tampil),
        "a_bytes": _to_bytes(doc_a),
        "a_name": f"Lampiran 5A - Informasi Uji Developer Site{akhiran}.docx",
        "b_bytes": _to_bytes(doc_b),
        "b_name": f"Lampiran 5B - Matriks Asesmen Uji Developer Site{akhiran}.docx",
    }

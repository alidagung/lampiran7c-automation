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

import copy
import io
import re
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
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

# Transaction Status Inquiry = Layanan API TERSENDIRI (baris paling akhir 5B).
# Di template asli baris ini berada di dalam grup Virtual Account; saat dibuat,
# baris itu dikeluarkan dari grup VA dan dijadikan baris sendiri.
LAYANAN_TSI = "API Transaction Status Inquiry"
SUB_API_TSI = "API Transaction Status Inquiry"
# Dianggap digunakan bila ada case dengan Nama Modul ini (section mana pun)
# yang statusnya bukan "Tidak dites".
MODUL_TSI = "transaction status inquiry"
BARIS_CONTOH_TUNGGAL = "API RTGS Transfer"   # baris tanpa sel gabungan untuk disalin

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


# Sub API -> (section UAT yang dicari, fungsi pencocok Nama Modul).
# section None = cari di SEMUA section. Nama Modul dibandingkan dalam huruf kecil.
def _modul_mengandung(*kata):
    return lambda modul: any(k in modul for k in kata)


def _modul_sama(*nilai):
    return lambda modul: modul in nilai


SUB_API_KE_MODUL = {
    "API Balance Inquiry": (("Balance Services",), _modul_mengandung("balance inquiry")),
    "API Internal Account Inquiry": (("Intrabank Transfer",), _modul_mengandung("internal account inquiry")),
    "API Intrabank Transfer": (("Intrabank Transfer",), _modul_mengandung("intrabank")),
    "API External Account Inquiry": (("Interbank Transfer",),
                                     _modul_mengandung("eksternal account inquiry", "external account inquiry")),
    "API Interbank Transfer": (("Interbank Transfer",), _modul_mengandung("interbank")),
    "API RTGS Transfer": (("RTGS Transfer",), _modul_mengandung("rtgs")),
    "API SKNBI Transfer": (("SKNBI Transfer",), _modul_mengandung("skn")),
    "API Virtual Account – Inquiry": (("Transfer VA",), _modul_sama("inquiry")),
    "API Virtual Account – Payment": (("Transfer VA",), _modul_sama("payment")),
    "API Virtual Account – Inquiry Status": (("Transfer VA",), lambda m: m.startswith("inquiry status")),
    SUB_API_TSI: (None, _modul_mengandung(MODUL_TSI)),
}

STATUS_DIGUNAKAN = "digunakan"
STATUS_TIDAK = "tidak digunakan"
STATUS_CEK = "perlu dicek"


def _status_aktif(hasil_aktual):
    """True bila case benar-benar dipakai (bukan 'Tidak dites')."""
    return str(hasil_aktual or "").strip().lower() not in ("", "tidak dites")


def validasi_api(uat_data):
    """
    Validasi Sub API apa saja yang DIGUNAKAN mitra, langsung dari UAT Script.

    Aturan per Sub API:
      - Cari case dengan Nama Modul yang cocok (lihat SUB_API_KE_MODUL).
      - "digunakan"       : minimal satu case cocok berstatus bukan "Tidak dites"
      - "tidak digunakan" : case cocok ada, tetapi semuanya "Tidak dites"
      - "perlu dicek"     : tidak ada case dengan Nama Modul yang cocok; status
                            diperkirakan dari section-nya (ada case aktif =
                            dianggap digunakan). Mohon dicek manual.

    Returns:
        list of dict (urut sesuai template 5B):
          {layanan_api, sub_api, status, digunakan (bool), bukti (str)}
    """
    hasil = []
    for info in daftar_sub_api():
        sub = info["sub_api"]
        sections, cocok = SUB_API_KE_MODUL.get(sub, (LAYANAN_KE_SECTION.get(info["layanan_api"], ()), None))
        nama_section = list(uat_data.keys()) if sections is None else list(sections)

        kasus = []
        if cocok is not None:
            for s in nama_section:
                for r in uat_data.get(s, []):
                    if cocok(str(r.get("nama_modul", "")).strip().lower()):
                        kasus.append(r)

        if kasus:
            aktif = [r for r in kasus if _status_aktif(r.get("hasil_aktual"))]
            if aktif:
                status = STATUS_DIGUNAKAN
                bukti = ", ".join(f"{r.get('nomor_kasus_tes', '?')} {str(r.get('hasil_aktual', '')).strip()}"
                                  for r in aktif[:4]) + (" …" if len(aktif) > 4 else "")
            else:
                status = STATUS_TIDAK
                bukti = f"{len(kasus)} case, semua Tidak dites"
        elif sections is None:
            # Sub API lintas-section (mis. Transaction Status Inquiry): tanpa
            # Nama Modul yang cocok tidak bisa diperkirakan -> tidak digunakan,
            # tetapi ditandai agar dicek manual.
            status, bukti = STATUS_CEK, "Nama Modul tidak ditemukan di UAT Script"
            info = {**info, "_aktif_section": False}
        else:
            sec_rows = [r for s in nama_section for r in uat_data.get(s, [])]
            if not sec_rows:
                status, bukti = STATUS_TIDAK, "tidak ada di UAT Script"
            else:
                status = STATUS_CEK
                aktif = any(_status_aktif(r.get("hasil_aktual")) for r in sec_rows)
                bukti = ("Nama Modul tidak dikenali; section "
                         + ("ada case aktif" if aktif else "semua Tidak dites"))
                info = {**info, "_aktif_section": aktif}
        digunakan = status == STATUS_DIGUNAKAN or (status == STATUS_CEK and info.get("_aktif_section", False))
        hasil.append({"layanan_api": info["layanan_api"], "sub_api": sub,
                      "status": status, "digunakan": digunakan, "bukti": bukti})
    return hasil


def sub_api_digunakan(uat_data):
    """Daftar nama Sub API yang digunakan mitra menurut UAT Script."""
    return [v["sub_api"] for v in validasi_api(uat_data) if v["digunakan"]]


def layanan_digunakan(uat_data):
    """
    Daftar Layanan API (urut sesuai template 5B) yang digunakan mitra menurut
    UAT Script = Layanan yang minimal satu Sub API-nya digunakan.

    Args:
        uat_data: dict hasil main.read_uat_script()
    """
    pakai = {v["layanan_api"] for v in validasi_api(uat_data) if v["digunakan"]}
    return [lay for lay in daftar_layanan() if lay in pakai]


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
    """
    Isi sel dengan teks, memakai format paragraf & run yang sudah ada.

    Teks ditulis ke paragraf PERTAMA YANG BERISI teks (bila ada), lalu paragraf
    lain dalam sel dihapus agar tidak tersisa baris kosong. Ini menangani sel
    template yang memiliki paragraf kosong sebelum/ sesudah isinya.
    """
    paras = cell.paragraphs
    target = next((p for p in paras if p.text.strip()), paras[0])
    runs = target.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r.text = ""
    else:
        target.add_run(text)
    for p in paras:
        if p is not target:
            p._p.getparent().remove(p._p)


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
# TATA LETAK: LANDSCAPE & TABEL PRESISI (Lampiran 5B / 6B)
# ============================================================

# Margin halaman landscape (Narrow 1,27 cm, sama dengan Lampiran 7C) agar tabel
# cukup lebar: lebar area tulis = 27,94 - 2 x 1,27 = 25,4 cm (14400 twips).
MARGIN_LANDSCAPE_EMU = 457200   # 1,27 cm

# Bobot lebar kolom grid (twips, dikonversi proporsional ke lebar area tulis).
# Diukur dengan metrik font Calibri 11pt agar teks terpanjang muat:
#   Sub API "API Virtual Account – Inquiry Status" 1 baris, email User DevSite
#   tidak terpotong di tengah kata, header "Response" (bold) muat 1 baris.
# 5B: No | Layanan API | Sub API | User DevSite | Pos Req | Pos Resp | Neg Req | Neg Resp
BOBOT_KOLOM_5B = (567, 2240, 3473, 3720, 1100, 1100, 1100, 1100)

TINGGI_HEADER_TWIPS = 454   # 0,8 cm (minimal) untuk 2 baris header
TINGGI_DATA_TWIPS = 567     # 1,0 cm (minimal) untuk baris data


def _set_landscape(doc, margin_emu=MARGIN_LANDSCAPE_EMU):
    """Ubah semua section menjadi LANDSCAPE (lebar & tinggi ditukar) + margin."""
    for s in doc.sections:
        if s.orientation != WD_ORIENT.LANDSCAPE:
            w, h = s.page_width, s.page_height
            s.orientation = WD_ORIENT.LANDSCAPE
            s.page_width, s.page_height = max(w, h), min(w, h)
        if margin_emu is not None:
            s.left_margin = s.right_margin = margin_emu
            s.top_margin = s.bottom_margin = margin_emu


def _lebar_area_twips(doc):
    s = doc.sections[0]
    return int(round((s.page_width - s.left_margin - s.right_margin) / 635))  # EMU -> twips


def _atur_elemen(parent, tag, attrs, sebelum=None):
    """Buat/ganti elemen anak `tag` pada parent dengan atribut attrs."""
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        if sebelum is not None and parent.find(qn(sebelum)) is not None:
            parent.find(qn(sebelum)).addprevious(el)
        else:
            parent.append(el)
    for k, v in attrs.items():
        el.set(qn(k), str(v))
    return el


def _rapikan_tabel(doc, table, bobot, baris_header=BARIS_DATA_MULAI):
    """
    Jadikan tabel PRESISI di halaman landscape:
      - lebar tabel = tepat lebar area tulis (tidak keluar margin), rata tengah
      - lebar kolom grid & sel konsisten (proporsional terhadap `bobot`)
      - tinggi baris seragam (header & data), baris tidak terpotong antar halaman
      - header berulang otomatis bila tabel berlanjut ke halaman berikutnya
      - isi sel rata tengah vertikal, spasi paragraf 0
    """
    tbl = table._tbl
    total = _lebar_area_twips(doc)
    jml = sum(bobot)
    lebar = [int(b * total / jml) for b in bobot]
    lebar[-1] += total - sum(lebar)          # sisa pembulatan ke kolom terakhir

    tblPr = tbl.tblPr
    _atur_elemen(tblPr, "w:tblW", {"w:w": total, "w:type": "dxa"})
    _atur_elemen(tblPr, "w:jc", {"w:val": "center"})
    _atur_elemen(tblPr, "w:tblInd", {"w:w": 0, "w:type": "dxa"})
    _atur_elemen(tblPr, "w:tblLayout", {"w:type": "fixed"})

    grid = tbl.tblGrid
    for gc in list(grid.findall(qn("w:gridCol"))):
        grid.remove(gc)
    for w in lebar:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        grid.append(gc)

    for ri, tr in enumerate(tbl.tr_lst):
        trPr = tr.get_or_add_trPr()
        for old in trPr.findall(qn("w:trHeight")):
            trPr.remove(old)
        h = TINGGI_HEADER_TWIPS if ri < baris_header else TINGGI_DATA_TWIPS
        _atur_elemen(trPr, "w:trHeight", {"w:val": h, "w:hRule": "atLeast"})
        _atur_elemen(trPr, "w:cantSplit", {})
        if ri < baris_header:
            _atur_elemen(trPr, "w:tblHeader", {})

        col = 0
        for tc in tr.tc_lst:
            tcPr = tc.get_or_add_tcPr()
            span_el = tcPr.find(qn("w:gridSpan"))
            span = int(span_el.get(qn("w:val"))) if span_el is not None else 1
            _atur_elemen(tcPr, "w:tcW", {"w:w": sum(lebar[col:col + span]), "w:type": "dxa"})
            _atur_elemen(tcPr, "w:vAlign", {"w:val": "center"})
            col += span
            for p in tc.iter(qn("w:p")):
                pPr = p.get_or_add_pPr()
                _atur_elemen(pPr, "w:spacing",
                             {"w:before": 0, "w:after": 0, "w:line": 240, "w:lineRule": "auto"})


# ============================================================
# NILAI DEFAULT (dibaca dari template)
# ============================================================

def _template_5b():
    """
    Muat template 5B lalu NORMALISASI: baris "API Transaction Status Inquiry"
    dikeluarkan dari grup Virtual Account dan ditambahkan sebagai baris
    tersendiri di akhir tabel (format disalin dari baris tanpa sel gabungan).
    """
    doc = Document(str(TEMPLATE_5B))
    table = doc.tables[0]
    trs = list(table._tbl.tr_lst)[BARIS_DATA_MULAI:]

    def teks(tc):
        return _Cell(tc, table).text.strip()

    tr_tsi = next((tr for tr in trs if teks(tr.tc_lst[2]) == SUB_API_TSI), None)
    tr_contoh = next((tr for tr in trs if teks(tr.tc_lst[1]) == BARIS_CONTOH_TUNGGAL), None)
    if tr_tsi is None or tr_contoh is None:
        return doc  # template berbeda: biarkan apa adanya

    baru = copy.deepcopy(tr_contoh)
    tr_tsi.getparent().remove(tr_tsi)
    table._tbl.append(baru)
    sel = baru.tc_lst
    _set_cell_text(_Cell(sel[1], table), LAYANAN_TSI)
    _set_cell_text(_Cell(sel[2], table), SUB_API_TSI)
    return doc


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
    t = _template_5b().tables[0]
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


def build_lampiran_5b(nomor_surat, user_devsite, tanda=None, layanan_tampil=None,
                      sub_api_tampil=None):
    """
    Args:
        nomor_surat : str
        user_devsite: str, email user DevSite (diisi ke semua baris)
        tanda       : list (per baris sub API) berisi 4 bool
                      [positif_req, positif_resp, negatif_req, negatif_resp];
                      True -> Ѵ, False -> X. None = semua Ѵ.
        layanan_tampil: daftar Layanan API yang ditampilkan (semua Sub API-nya).
        sub_api_tampil: daftar Sub API yang ditampilkan (lebih spesifik; bila
                      diisi, mengalahkan layanan_tampil). Keduanya None = semua.
    """
    doc = _template_5b()

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

    # Saring Sub API (None = semua) & nomori ulang kolom No
    rows = daftar_sub_api()
    saring_sub_api(t, rows, pilih_sub_api(rows, layanan_tampil, sub_api_tampil))

    # Halaman landscape + tabel presisi
    _set_landscape(doc)
    _rapikan_tabel(doc, t, BOBOT_KOLOM_5B)
    return doc


def _vmerge(tc):
    tcPr = tc.tcPr
    el = tcPr.find(qn("w:vMerge")) if tcPr is not None else None
    if el is None:
        return None
    return el.get(qn("w:val")) or "continue"


def _set_vmerge(tc, nilai):
    """nilai: 'restart' / 'continue' / None (hapus gabungan)."""
    tcPr = tc.get_or_add_tcPr()
    el = tcPr.find(qn("w:vMerge"))
    if nilai is None:
        if el is not None:
            tcPr.remove(el)
        return
    if el is None:
        el = OxmlElement("w:vMerge")
        tcPr.append(el)
    if nilai == "restart":
        el.set(qn("w:val"), "restart")
    elif qn("w:val") in el.attrib:
        del el.attrib[qn("w:val")]


def saring_sub_api(table, rows_info, tampil_sub):
    """
    Tampilkan hanya Sub API dalam `tampil_sub` (set nama Sub API).

    Baris Sub API lain dihapus. Sel gabungan vertikal (No, Layanan API, User,
    dsb.) disusun ulang agar tetap rapi walau baris pertama grup terhapus.
    Kolom No dinomori ulang 1, 2, 3, ... per Layanan API.

    Args:
        table     : tabel matriks (5B / 6B)
        rows_info : list dict {layanan_api, sub_api} sejajar baris data template
        tampil_sub: set nama Sub API yang ditampilkan
    """
    trs = list(table._tbl.tr_lst)[BARIS_DATA_MULAI:]
    grup = []
    for info, tr in zip(rows_info, trs):
        if grup and grup[-1][0] == info["layanan_api"]:
            grup[-1][1].append((info, tr))
        else:
            grup.append((info["layanan_api"], [(info, tr)]))

    nomor = 0
    for _layanan, items in grup:
        first_tr = items[0][1]
        kolom_gabung = [ci for ci, tc in enumerate(first_tr.tc_lst) if _vmerge(tc) == "restart"]
        sumber = {ci: copy.deepcopy(first_tr.tc_lst[ci]) for ci in kolom_gabung}
        simpan = [tr for info, tr in items if info["sub_api"] in tampil_sub]
        for _info, tr in items:
            if tr not in simpan:
                tr.getparent().remove(tr)
        if not simpan:
            continue

        baru_pertama = simpan[0]
        if baru_pertama is not first_tr:
            # Pindahkan isi sel gabungan (No/Layanan/User/...) ke baris pertama baru
            for ci in kolom_gabung:
                lama = baru_pertama.tc_lst[ci]
                pengganti = copy.deepcopy(sumber[ci])
                lama.addprevious(pengganti)
                lama.getparent().remove(lama)
        for ci in kolom_gabung:
            for k, tr in enumerate(simpan):
                tc = tr.tc_lst[ci]
                if len(simpan) == 1:
                    _set_vmerge(tc, None)
                else:
                    _set_vmerge(tc, "restart" if k == 0 else "continue")

        nomor += 1
        _set_cell_text(_Cell(baru_pertama.tc_lst[0], table), str(nomor))


def pilih_sub_api(rows_info, layanan_tampil=None, sub_api_tampil=None):
    """Set Sub API yang ditampilkan (sub_api_tampil > layanan_tampil > semua)."""
    if sub_api_tampil is not None:
        return set(sub_api_tampil)
    if layanan_tampil is not None:
        return _sub_dari_layanan(rows_info, layanan_tampil)
    return {r["sub_api"] for r in rows_info}


def _sub_dari_layanan(rows_info, layanan_tampil):
    return {r["sub_api"] for r in rows_info if r["layanan_api"] in set(layanan_tampil)}


def _saring_layanan(table, tampil):
    """Kompatibilitas: saring per Layanan API (semua Sub API-nya)."""
    rows = daftar_sub_api()
    saring_sub_api(table, rows, _sub_dari_layanan(rows, tampil))


def baca_uat(source):
    """read_uat_script() untuk path/bytes/file-like."""
    import main  # impor lokal (main tidak bergantung pada modul ini)
    if isinstance(source, (bytes, bytearray)):
        src = io.BytesIO(source)
    else:
        src = source
        if hasattr(src, "seek"):
            src.seek(0)
    return main.read_uat_script(src)


def resolve_sub_api(source, layanan_tampil=None, sub_api_tampil=None):
    """
    Tentukan Sub API yang ditampilkan: pilihan manual bila ada, bila tidak
    hasil validasi UAT Script. Bila hasilnya kosong, tampilkan semua (tabel
    tidak boleh kosong).
    """
    rows = daftar_sub_api()
    if sub_api_tampil is not None or layanan_tampil is not None:
        pilih = pilih_sub_api(rows, layanan_tampil, sub_api_tampil)
    else:
        pilih = set(sub_api_digunakan(baca_uat(source)))
    return pilih or {r["sub_api"] for r in rows}


def layanan_dari_sub(sub_tampil):
    out = []
    for r in daftar_sub_api():
        if r["sub_api"] in sub_tampil and r["layanan_api"] not in out:
            out.append(r["layanan_api"])
    return out


def _to_bytes(doc):
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def nama_file_aman(teks):
    return re.sub(r'[\\/:*?"<>|]+', "-", teks).strip()


def convert_lampiran_5_bytes(source, product_key, penandatangan, user_devsite,
                             tanda=None, tanggal=None, nomor_surat=None,
                             layanan_tampil=None, sub_api_tampil=None):
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
        layanan_tampil: Layanan API yang ditampilkan di 5B (semua Sub API-nya)
        sub_api_tampil: Sub API yang ditampilkan di 5B (lebih spesifik).
                       Keduanya None -> hasil validasi_api() dari UAT Script.

    Returns:
        dict {nomor_surat, tanggal, layanan, sub_api, a_bytes, a_name, b_bytes, b_name}
        atau None bila produk belum didukung.
    """
    if product_key not in PRODUK_DIDUKUNG:
        return None
    tanggal = tanggal or date.today()
    nomor = (nomor_surat or "").strip() or read_nomor_referensi(source)

    sub_tampil = resolve_sub_api(source, layanan_tampil, sub_api_tampil)

    doc_a = build_lampiran_5a(nomor, tanggal, penandatangan)
    doc_b = build_lampiran_5b(nomor, user_devsite, tanda, sub_api_tampil=sub_tampil)
    akhiran = f" - {nama_file_aman(nomor)}" if nomor else ""
    return {
        "nomor_surat": nomor,
        "tanggal": tanggal,
        "layanan": layanan_dari_sub(sub_tampil),
        "sub_api": [r["sub_api"] for r in daftar_sub_api() if r["sub_api"] in sub_tampil],
        "a_bytes": _to_bytes(doc_a),
        "a_name": f"Lampiran 5A - Informasi Uji Developer Site{akhiran}.docx",
        "b_bytes": _to_bytes(doc_b),
        "b_name": f"Lampiran 5B - Matriks Asesmen Uji Developer Site{akhiran}.docx",
    }

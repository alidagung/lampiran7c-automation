"""
Penggabung dokumen Word (Lampiran 5A ... 7C menjadi SATU file)
==============================================================
Setiap dokumen sumber menjadi satu SECTION tersendiri di dokumen gabungan,
sehingga yang berikut ini tetap SAMA seperti dokumen aslinya:

  - ukuran & orientasi halaman, margin (portrait / landscape per lampiran)
  - header & footer (mis. logo BSS di 5A/5B tetap ada; lampiran tanpa header
    tetap tanpa header, tidak "mewarisi" header lampiran sebelumnya)
  - format teks: style tiap dokumen disalin dengan nama unik (awalan per
    dokumen), sehingga mis. Normal 9pt milik Lampiran 7C tidak tertimpa
    Normal 11pt milik template 5A
  - gambar & hyperlink (relationship disalin ulang)

Setiap lampiran dimulai di halaman baru (section break "next page").
"""

import copy
import io
import re

from lxml import etree
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_ATTR_REL = (f"{{{_R}}}embed", f"{{{_R}}}id", f"{{{_R}}}link")
_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_TAG_REF_HF = (qn("w:headerReference"), qn("w:footerReference"))


def _as_doc(x):
    if isinstance(x, (bytes, bytearray)):
        return Document(io.BytesIO(x))
    return x


# ------------------------------------------------------------
# Style
# ------------------------------------------------------------

def _styles(doc):
    return {s.get(qn("w:styleId")): s for s in doc.styles.element.findall(qn("w:style"))}


def _default_style_id(doc, jenis):
    for s in doc.styles.element.findall(qn("w:style")):
        if s.get(qn("w:type")) == jenis and s.get(qn("w:default")) in ("1", "true", "on"):
            return s.get(qn("w:styleId"))
    return None


def _defaults_xml(doc):
    dd = doc.styles.element.find(qn("w:docDefaults"))
    return b"" if dd is None else etree.tostring(dd)


def _sisip_basedon(style_el, nilai):
    """Sisipkan <w:basedOn> pada posisi yang sah (setelah w:name / w:aliases)."""
    b = OxmlElement("w:basedOn")
    b.set(qn("w:val"), nilai)
    jangkar = style_el.find(qn("w:aliases"))
    if jangkar is None:
        jangkar = style_el.find(qn("w:name"))
    if jangkar is not None:
        jangkar.addnext(b)
    else:
        style_el.insert(0, b)


def _ganti_style_ref(elements, prefix, def_para, src_styles):
    """
    Ubah referensi style di `elements` menjadi id berawalan `prefix`.
    Paragraf tanpa pStyle diberi pStyle = style paragraf default sumber
    (agar tidak memakai style default dokumen gabungan).
    Mengembalikan set id style sumber yang dipakai.
    """
    dipakai = set()
    for el in elements:
        for p in el.iter(qn("w:p")):
            pPr = p.find(qn("w:pPr"))
            ps = pPr.find(qn("w:pStyle")) if pPr is not None else None
            if ps is not None:
                sid = ps.get(qn("w:val"))
                if sid not in src_styles:
                    sid = def_para
                if sid:
                    ps.set(qn("w:val"), prefix + sid)
                    dipakai.add(sid)
            elif def_para:
                if pPr is None:
                    pPr = OxmlElement("w:pPr")
                    p.insert(0, pPr)
                ps = OxmlElement("w:pStyle")
                ps.set(qn("w:val"), prefix + def_para)
                pPr.insert(0, ps)
                dipakai.add(def_para)
        for tag in ("w:rStyle", "w:tblStyle"):
            for e in el.iter(qn(tag)):
                sid = e.get(qn("w:val"))
                if sid in src_styles:
                    e.set(qn("w:val"), prefix + sid)
                    dipakai.add(sid)
    return dipakai


def _salin_styles(master, src, prefix, label, dipakai):
    """Salin style sumber yang dipakai (+ rantai basedOn) ke master dengan id unik."""
    src_styles = _styles(src)
    sudah = set(_styles(master))
    perlu, antre = set(), list(dipakai)
    while antre:
        sid = antre.pop()
        if sid in perlu or sid not in src_styles:
            continue
        perlu.add(sid)
        b = src_styles[sid].find(qn("w:basedOn"))
        if b is not None:
            antre.append(b.get(qn("w:val")))

    # Bila docDefaults sumber berbeda dari master, buat style dasar berisi
    # docDefaults sumber sebagai akar rantai style paragraf.
    base_id = None
    if _defaults_xml(src) != _defaults_xml(master):
        base_id = prefix + "_DocDefaults"
        if base_id not in sudah:
            st = OxmlElement("w:style")
            st.set(qn("w:type"), "paragraph")
            st.set(qn("w:styleId"), base_id)
            nm = OxmlElement("w:name")
            nm.set(qn("w:val"), f"{label} (default)")
            st.append(nm)
            st.append(OxmlElement("w:semiHidden"))
            dd = src.styles.element.find(qn("w:docDefaults"))
            ppr_d = dd.find(qn("w:pPrDefault")) if dd is not None else None
            rpr_d = dd.find(qn("w:rPrDefault")) if dd is not None else None
            if ppr_d is not None and ppr_d.find(qn("w:pPr")) is not None:
                st.append(copy.deepcopy(ppr_d.find(qn("w:pPr"))))
            if rpr_d is not None and rpr_d.find(qn("w:rPr")) is not None:
                st.append(copy.deepcopy(rpr_d.find(qn("w:rPr"))))
            master.styles.element.append(st)

    for sid in perlu:
        new_id = prefix + sid
        if new_id in sudah:
            continue
        el = copy.deepcopy(src_styles[sid])
        el.set(qn("w:styleId"), new_id)
        el.attrib.pop(qn("w:default"), None)
        nm = el.find(qn("w:name"))
        if nm is not None:
            nm.set(qn("w:val"), f"{label} {nm.get(qn('w:val'))}")
        b = el.find(qn("w:basedOn"))
        if b is not None:
            b.set(qn("w:val"), prefix + b.get(qn("w:val")))
        elif base_id and el.get(qn("w:type")) == "paragraph":
            _sisip_basedon(el, base_id)
        for tag in ("w:link", "w:next"):
            e = el.find(qn(tag))
            if e is not None:
                if e.get(qn("w:val")) in perlu:
                    e.set(qn("w:val"), prefix + e.get(qn("w:val")))
                else:
                    el.remove(e)
        # Numbering milik dokumen sumber tidak ikut disalin -> lepas numPr
        for numpr in el.iter(qn("w:numPr")):
            numpr.getparent().remove(numpr)
        master.styles.element.append(el)


# ------------------------------------------------------------
# Relationship, id unik, dsb.
# ------------------------------------------------------------

def _remap_rel(elements, src_part, dst_part):
    """Salin relationship (gambar/hyperlink) yang dirujuk elemen ke dst_part."""
    for el in elements:
        for node in el.iter():
            for attr in _ATTR_REL:
                rid = node.get(attr)
                if not rid or rid not in src_part.rels:
                    continue
                rel = src_part.rels[rid]
                if rel.is_external:
                    new = dst_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                elif rel.reltype == RT.IMAGE:
                    new, _img = dst_part.get_or_add_image(io.BytesIO(rel.target_part.blob))
                else:
                    raise ValueError(f"Relationship tidak didukung saat menggabung: {rel.reltype}")
                node.set(attr, new)


class _Penomor:
    """Membuat id unik untuk bookmark & gambar di seluruh dokumen gabungan."""

    def __init__(self):
        self.n = 10000

    def berikut(self):
        self.n += 1
        return str(self.n)


def _unikkan(elements, prefix, penomor):
    peta_bm = {}
    for el in elements:
        for node in el.iter():
            # paraId/textId Word 2010 bersifat opsional; dibuang agar tak dobel
            node.attrib.pop(f"{{{_W14}}}paraId", None)
            node.attrib.pop(f"{{{_W14}}}textId", None)
            if node.tag in (qn("w:bookmarkStart"), qn("w:bookmarkEnd")):
                lama = node.get(qn("w:id"))
                peta_bm.setdefault(lama, penomor.berikut())
                node.set(qn("w:id"), peta_bm[lama])
                if node.tag == qn("w:bookmarkStart") and node.get(qn("w:name")):
                    node.set(qn("w:name"), f"{prefix}{node.get(qn('w:name'))}"[:40])
            elif node.tag == f"{{{_WP}}}docPr":
                node.set("id", penomor.berikut())


# ------------------------------------------------------------
# Section
# ------------------------------------------------------------

def _tanpa_header_footer(sectPr):
    s = copy.deepcopy(sectPr)
    for c in list(s):
        if c.tag in _TAG_REF_HF:
            s.remove(c)
    return s


def _tutup_section_master(master):
    """Pindahkan sectPr akhir body master ke paragraf terakhir (section break)."""
    body = master.element.body
    sect = body.find(qn("w:sectPr"))
    terakhir = sect.getprevious()
    if terakhir is not None and terakhir.tag == qn("w:p"):
        pPr = terakhir.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            terakhir.insert(0, pPr)
        if pPr.find(qn("w:sectPr")) is None:
            chg = pPr.find(qn("w:pPrChange"))
            salinan = copy.deepcopy(sect)
            if chg is not None:
                chg.addprevious(salinan)
            else:
                pPr.append(salinan)
            return sect
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pPr.append(copy.deepcopy(sect))
    p.append(pPr)
    sect.addprevious(p)
    return sect


def _salin_header_footer(src_hf, dst_hf, prefix, src_doc, master, label, penomor):
    """Isi header/footer dst dengan salinan isi src (style & gambar ikut)."""
    dst_hf.is_linked_to_previous = False
    dst_el = dst_hf._element
    for c in list(dst_el):
        dst_el.remove(c)
    if src_hf is None or src_hf.is_linked_to_previous:
        dst_el.append(OxmlElement("w:p"))       # header kosong (tidak mewarisi)
        return
    isi = [copy.deepcopy(c) for c in src_hf._element]
    src_styles = _styles(src_doc)
    dipakai = _ganti_style_ref(isi, prefix, _default_style_id(src_doc, "paragraph"), src_styles)
    _salin_styles(master, src_doc, prefix, label, dipakai)
    _remap_rel(isi, src_hf.part, dst_hf.part)
    _unikkan(isi, prefix, penomor)
    for c in isi:
        dst_el.append(c)
    if not len(dst_el):
        dst_el.append(OxmlElement("w:p"))


# ------------------------------------------------------------
# API
# ------------------------------------------------------------

def gabung_dokumen(daftar, label=None):
    """
    Gabungkan dokumen (bytes atau docx.Document) menjadi satu Document.

    Args:
        daftar: list dokumen, urut sesuai urutan di hasil gabungan
        label : list label per dokumen (dipakai sebagai awalan nama style,
                mis. "5A"); default "D1", "D2", ...

    Returns:
        docx.Document gabungan
    """
    docs = [_as_doc(d) for d in daftar]
    if not docs:
        raise ValueError("Tidak ada dokumen untuk digabung.")
    label = list(label) if label else [f"D{i + 1}" for i in range(len(docs))]
    master = docs[0]
    penomor = _Penomor()
    jumlah_section = [len(master.sections)]

    for i, src in enumerate(docs[1:], start=1):
        lab = re.sub(r"[^A-Za-z0-9]", "", label[i]) or f"D{i + 1}"
        prefix = f"L{lab}_"
        sect_master = _tutup_section_master(master)

        body = src.element.body
        sect_src = body.find(qn("w:sectPr"))
        elemen = [e for e in body if e is not sect_src]
        src_styles = _styles(src)
        dipakai = _ganti_style_ref(elemen, prefix, _default_style_id(src, "paragraph"), src_styles)
        _salin_styles(master, src, prefix, f"Lampiran {label[i]}", dipakai)
        _remap_rel(elemen, src.part, master.part)
        _unikkan(elemen, prefix, penomor)

        for e in elemen:
            sect_master.addprevious(e)
        sect_master.getparent().replace(sect_master, _tanpa_header_footer(sect_src))
        jumlah_section.append(len(src.sections))

    # Header/footer per section: section pertama (dokumen master) dibiarkan;
    # section dokumen berikutnya diisi salinan header/footer aslinya (atau
    # dikosongkan bila aslinya tanpa header/footer).
    sec_master = master.sections
    idx = jumlah_section[0]
    for i, src in enumerate(docs[1:], start=1):
        lab = re.sub(r"[^A-Za-z0-9]", "", label[i]) or f"D{i + 1}"
        prefix = f"L{lab}_"
        for k, s_src in enumerate(src.sections):
            s_dst = sec_master[idx + k]
            for jenis in ("header", "footer"):
                _salin_header_footer(getattr(s_src, jenis), getattr(s_dst, jenis),
                                     prefix, src, master, f"Lampiran {label[i]}", penomor)
            if s_src.different_first_page_header_footer:
                s_dst.different_first_page_header_footer = True
                for jenis in ("first_page_header", "first_page_footer"):
                    _salin_header_footer(getattr(s_src, jenis), getattr(s_dst, jenis),
                                         prefix, src, master, f"Lampiran {label[i]}", penomor)
        idx += len(src.sections)
    return master


def gabung_bytes(daftar, label=None):
    """Seperti gabung_dokumen(), mengembalikan bytes .docx."""
    doc = gabung_dokumen(daftar, label)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

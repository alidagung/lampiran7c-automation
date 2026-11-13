"""
Aplikasi Web Automation Dokumen ASPI
====================================
Aplikasi web berbasis Streamlit untuk mengubah file UAT Script (Excel) menjadi
dokumen ASPI (Word): Lampiran 7C, UAT Result, Lampiran 5A/5B, 6A/6B, dan 7A/7B.

Alur (3 langkah):
    1. Unggah UAT Script  -> info mitra (nama PT, nomor referensi) langsung tampil
    2. Validasi API       -> Sub API yang digunakan dideteksi otomatis dari UAT
                             Script (bisa dikoreksi), isi data mitra
    3. Buat & unduh       -> satu klik membuat semua dokumen, unduh satu per
                             satu atau sekaligus (.zip)

Cara pakai:
    pip install -r requirements.txt
    streamlit run app.py
"""

import base64
import html
import io
import json
import zipfile
from datetime import date, datetime
from pathlib import Path

import streamlit as st

from main import (
    convert_uat_to_lampiran_bytes,
    convert_uat_to_result_bytes,
    detect_product,
    format_tanggal_7c,
    read_uat_script,
    PRODUCT_PROFILES,
    OUTPUT_FILENAME,
    LOGO_BSS,
)
import lampiran5
import lampiran6
import lampiran7

# ============================================================
# KONSTANTA & PENGATURAN LOKAL
# ============================================================

PRODUCT_OPTIONS = {
    "Deteksi otomatis": "AUTO",
    "Fund Transfer & Virtual Account (VA/FT)": "FT_VA",
    "QRIS": "QRIS",
}
PRODUCT_LABEL = {"FT_VA": "VA / Fund Transfer", "QRIS": "QRIS"}

BRAND = "#ED7D31"
BRAND_DARK = "#C55A11"
NAVY = "#1E2A47"
NAVY_2 = "#2B3A5E"
MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

STATUS_LABEL = {
    lampiran5.STATUS_DIGUNAKAN: "✅ Digunakan",
    lampiran5.STATUS_TIDAK: "⛔ Tidak digunakan",
    lampiran5.STATUS_CEK: "⚠️ Perlu dicek",
}

# Pengaturan yang jarang berubah (penandatangan, user DevSite bank) disimpan di
# komputer pengguna (tidak ikut ke GitHub) agar tidak perlu diketik ulang.
PENGATURAN_FILE = Path(__file__).resolve().parent / "pengaturan_lampiran5.json"


def muat_pengaturan():
    data = {**lampiran5.default_penandatangan(),
            "user_devsite": lampiran5.default_user_devsite()}
    try:
        if PENGATURAN_FILE.exists():
            data.update(json.loads(PENGATURAN_FILE.read_text(encoding="utf-8")))
    except Exception:  # noqa: BLE001 - file rusak -> pakai default
        pass
    return data


@st.cache_data(show_spinner=False)
def logo_data_uri():
    """Logo BSS sebagai data URI (untuk HTML); "" bila file logo tidak ada."""
    try:
        return "data:image/png;base64," + base64.b64encode(LOGO_BSS.read_bytes()).decode()
    except Exception:  # noqa: BLE001
        return ""


def logo_img(tinggi_px, kelas=""):
    uri = logo_data_uri()
    if not uri:
        return '<span style="font-size:1.6rem">🏦</span>'
    return f'<img src="{uri}" class="{kelas}" style="height:{tinggi_px}px" alt="Bank Sahabat Sampoerna">'


def simpan_pengaturan(data):
    PENGATURAN_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


@st.cache_data(show_spinner=False)
def analisa_uat(file_bytes):
    """
    Baca UAT Script sekali (di-cache) untuk informasi mitra & validasi API.
    Tidak mengubah file apa pun.
    """
    hasil = {"ok": True, "error": "", "produk": None, "nomor": "", "pengguna": "",
             "validasi": [], "jumlah_case": 0, "jumlah_section": 0}
    try:
        uat = read_uat_script(io.BytesIO(file_bytes))
        hasil["produk"] = detect_product(uat)
        hasil["validasi"] = lampiran5.validasi_api(uat)
        hasil["jumlah_case"] = sum(len(v) for v in uat.values())
        hasil["jumlah_section"] = len(uat)
    except Exception as e:  # noqa: BLE001
        hasil.update(ok=False, error=str(e))
    for kunci, fungsi in (("nomor", lampiran5.read_nomor_referensi),
                          ("pengguna", lampiran6.read_nama_pengguna)):
        try:
            hasil[kunci] = fungsi(io.BytesIO(file_bytes))
        except Exception:  # noqa: BLE001
            pass
    return hasil


# ============================================================
# HALAMAN & GAYA
# ============================================================

st.set_page_config(
    page_title="Automation Dokumen ASPI",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    f"""
    <style>
    /* ---------- Dasar ---------- */
    .stApp {{
        background:
            radial-gradient(1200px 380px at 85% -120px, rgba(237,125,49,0.16), transparent 70%),
            radial-gradient(900px 300px at -10% -80px, rgba(43,58,94,0.10), transparent 70%),
            #F5F6FA;
    }}
    .block-container {{ padding-top: 1.2rem; max-width: 1200px; }}

    /* ---------- Sidebar gelap ---------- */
    section[data-testid="stSidebar"] > div:first-child {{
        background: linear-gradient(180deg, {NAVY} 0%, {NAVY_2} 100%);
    }}
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] li,
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3,
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h4,
    section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
    section[data-testid="stSidebar"] [data-testid="stExpander"] summary p {{
        color: #EEF1F7 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{ color: #AEB8CC !important; }}
    section[data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,0.14); }}
    section[data-testid="stSidebar"] [data-testid="stExpander"] {{
        background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.12); border-radius: 12px;
    }}
    .sb-brand {{ display:flex; align-items:center; gap:10px; margin: 2px 0 2px 0; }}
    .sb-logo {{
        width: 38px; height: 38px; border-radius: 11px; display:flex; align-items:center; justify-content:center;
        background: linear-gradient(135deg, {BRAND}, {BRAND_DARK}); font-size: 20px;
        box-shadow: 0 6px 14px rgba(237,125,49,0.35);
    }}
    .sb-title {{ color:#fff; font-weight: 800; font-size: 1.02rem; line-height: 1.15; }}
    .sb-sub {{ color:#AEB8CC; font-size: 0.78rem; }}
    .sb-section {{
        color: {BRAND} !important; font-weight: 800; font-size: 0.78rem; letter-spacing: .8px;
        text-transform: uppercase; margin: 6px 0 4px 0;
    }}
    .sb-box {{
        background: rgba(237,125,49,0.10); border: 1px solid rgba(237,125,49,0.35);
        border-radius: 12px; padding: 8px 12px; margin-bottom: 6px;
    }}
    .sb-box .k {{ color:#AEB8CC; font-size: 0.72rem; text-transform: uppercase; letter-spacing:.4px; font-weight:700; }}
    .sb-box .v {{ color:#fff; font-weight: 700; font-size: 0.92rem; word-break: break-word; }}

    /* ---------- Hero ---------- */
    .hero {{
        position: relative; overflow: hidden;
        background: linear-gradient(115deg, {NAVY} 0%, {NAVY_2} 45%, {BRAND_DARK} 100%);
        padding: 26px 32px; border-radius: 20px; color: #fff; margin-bottom: 14px;
        box-shadow: 0 14px 34px rgba(30,42,71,0.28);
        display: flex; align-items: center; gap: 18px;
    }}
    .hero::after {{
        content:""; position:absolute; right:-60px; top:-60px; width:240px; height:240px; border-radius:50%;
        background: radial-gradient(circle, rgba(237,125,49,0.55), rgba(237,125,49,0) 70%);
    }}
    .hero-logo {{
        width: 58px; height: 58px; border-radius: 16px; flex-shrink: 0; z-index: 1;
        background: linear-gradient(135deg, {BRAND}, {BRAND_DARK});
        display:flex; align-items:center; justify-content:center; font-size: 30px;
        box-shadow: 0 8px 20px rgba(237,125,49,0.45);
    }}
    .hero-teks {{ z-index: 1; }}
    .hero h1 {{ color:#fff; font-size: 1.65rem; margin: 0 0 3px 0; font-weight: 800; letter-spacing:-0.4px; }}
    .hero p  {{ color:#D7DEEA; font-size: 0.95rem; margin: 0; }}
    .hero-tag {{
        display:inline-block; margin: 8px 6px 0 0; padding: 3px 11px; border-radius: 999px;
        background: rgba(255,255,255,0.12); border: 1px solid rgba(255,255,255,0.22);
        font-size: 0.76rem; font-weight: 700; color: #fff;
    }}

    /* ---------- Kartu langkah (st.container(border=True)) ---------- */
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        background: #FFFFFF; border-radius: 18px !important;
        border: 1px solid rgba(30,42,71,0.08) !important;
        box-shadow: 0 6px 22px rgba(30,42,71,0.06);
    }}

    /* ---------- Stepper ---------- */
    .stepper {{ display:flex; gap: 10px; margin: 0 0 14px 0; flex-wrap: wrap; }}
    .step {{
        flex: 1; min-width: 180px; display:flex; align-items:center; gap:10px;
        padding: 11px 14px; border-radius: 14px; font-weight: 700; font-size: 0.9rem;
        background: #FFFFFF; color: #8A93A6; border: 1px solid rgba(30,42,71,0.08);
        box-shadow: 0 3px 10px rgba(30,42,71,0.04);
    }}
    .step .num {{
        width: 28px; height: 28px; border-radius: 50%; display:flex; align-items:center;
        justify-content:center; font-size: 0.84rem; font-weight: 800;
        background: #E3E7EF; color: #8A93A6; flex-shrink: 0;
    }}
    .step.done {{ color: #1E7E38; border-color: rgba(46,158,79,0.30); }}
    .step.done .num {{ background: #2E9E4F; color:#fff; }}
    .step.active {{ color: {BRAND_DARK}; border-color: rgba(237,125,49,0.45);
                    box-shadow: 0 6px 16px rgba(237,125,49,0.18); }}
    .step.active .num {{ background: linear-gradient(135deg, {BRAND}, {BRAND_DARK}); color:#fff; }}

    .judul-langkah {{ display:flex; align-items:center; gap:10px; font-size: 1.12rem; font-weight: 800;
                      color: {NAVY}; margin: 0 0 2px 0; }}
    .judul-langkah .no {{
        width: 28px; height: 28px; border-radius: 9px; display:flex; align-items:center; justify-content:center;
        background: linear-gradient(135deg, {BRAND}, {BRAND_DARK}); color:#fff; font-size: 0.86rem;
    }}
    .sub-langkah {{ color: #6B7280; font-size: 0.88rem; margin: 0 0 12px 38px; }}

    /* ---------- Chip info mitra ---------- */
    .chips {{ display:grid; grid-template-columns: repeat(4, minmax(0,1fr)); gap: 10px; margin-top: 8px; }}
    .chip {{ background: linear-gradient(180deg,#FFF8F2,#FFFFFF); border:1px solid rgba(237,125,49,0.22);
             border-radius: 14px; padding: 10px 14px; }}
    .chip .k {{ font-size: 0.72rem; color:#8A93A6; font-weight:800; text-transform: uppercase; letter-spacing: .4px; }}
    .chip .v {{ font-size: 0.98rem; font-weight: 800; color:{NAVY}; word-break: break-word; }}
    .chip .v.kosong {{ color:#C62828; }}

    /* ---------- Ringkasan validasi ---------- */
    .pill {{ display:inline-block; padding: 5px 13px; border-radius: 999px; font-size: 0.84rem;
             font-weight: 800; margin: 0 6px 8px 0; }}
    .pill-ok {{ background:#E6F6EB; color:#1E7E38; }}
    .pill-no {{ background:#EEF0F4; color:#6B7280; }}
    .pill-cek {{ background:#FFF1DC; color:#B26A00; }}

    /* ---------- Kartu statistik ---------- */
    .stat {{ border-radius: 16px; padding: 14px 16px; color:#fff; box-shadow: 0 8px 20px rgba(30,42,71,0.12); }}
    .stat .lbl {{ font-size: 0.8rem; opacity: 0.92; font-weight:700; }}
    .stat .val {{ font-size: 1.8rem; font-weight: 800; line-height:1.1; }}
    .stat-pass {{ background: linear-gradient(135deg,#34B35A,#1E7E38); }}
    .stat-fail {{ background: linear-gradient(135deg,#EF5A5A,#C62828); }}
    .stat-na   {{ background: linear-gradient(135deg,#F5B347,#D9871F); }}
    .stat-tot  {{ background: linear-gradient(135deg,{NAVY_2},{NAVY}); }}

    /* ---------- Kartu dokumen ---------- */
    .doc-ikon {{
        width: 40px; height: 40px; border-radius: 12px; display:flex; align-items:center; justify-content:center;
        background: #FFF3EA; font-size: 1.3rem;
    }}
    .doc-judul {{ font-weight: 800; font-size: 1.0rem; margin-top: 8px; color:{NAVY}; }}
    .doc-ket {{ color:#6B7280; font-size: 0.84rem; min-height: 2.5em; }}

    /* ---------- Tombol ---------- */
    .stButton>button, .stDownloadButton>button {{
        border-radius: 12px; font-weight: 800; transition: transform .08s ease, box-shadow .15s ease;
    }}
    .stButton>button[kind="primary"], .stDownloadButton>button[kind="primary"] {{
        background: linear-gradient(135deg, {BRAND}, {BRAND_DARK}); border: 0;
        box-shadow: 0 8px 18px rgba(237,125,49,0.32);
    }}
    .stButton>button:hover, .stDownloadButton>button:hover {{ transform: translateY(-1px); }}

    /* ---------- Unggah ---------- */
    section[data-testid="stFileUploaderDropzone"] {{
        border: 2px dashed rgba(237,125,49,0.65); border-radius: 16px;
        background: linear-gradient(180deg,#FFF8F2,#FFFFFF);
    }}

    .tgl-pratinjau {{
        background: #FFF8F2; border: 1px solid rgba(237,125,49,0.28); border-radius: 12px;
        padding: 8px 14px; font-size: 0.86rem; color: {NAVY}; margin: 2px 0 6px 0;
    }}

    /* ---------- Logo BSS ---------- */
    .sb-logo-card {{
        background: #FFFFFF; border-radius: 16px; padding: 12px 14px 10px 14px; text-align: center;
        box-shadow: 0 10px 24px rgba(0,0,0,0.25); margin-bottom: 8px;
    }}
    .sb-logo-card img {{ max-width: 100%; }}
    .sb-app {{ color:#FFFFFF; font-weight: 800; font-size: 0.98rem; text-align:center; margin-top: 6px; }}
    .sb-app-sub {{ color:#AEB8CC; font-size: 0.76rem; text-align:center; }}
    .hero-logo-card {{
        background: #FFFFFF; border-radius: 16px; padding: 10px 16px; flex-shrink: 0; z-index: 1;
        box-shadow: 0 10px 26px rgba(0,0,0,0.25); display:flex; align-items:center;
    }}
    .hero {{ background-size: 200% 200%; animation: geser 14s ease infinite; }}
    @keyframes geser {{ 0% {{background-position: 0% 50%}} 50% {{background-position: 100% 50%}}
                        100% {{background-position: 0% 50%}} }}
    .hero-kanan {{ margin-left: auto; z-index: 1; text-align: right; color: #D7DEEA; font-size: 0.8rem; }}
    .hero-kanan b {{ display:block; color:#fff; font-size: 1.5rem; line-height: 1.1; }}

    /* ---------- Kartu unduh utama ---------- */
    .unduh-kartu {{
        border-radius: 18px; padding: 18px 20px; color:#fff; margin-bottom: 8px;
        box-shadow: 0 12px 28px rgba(30,42,71,0.18); min-height: 150px;
    }}
    .unduh-lampiran {{ background: linear-gradient(135deg, {BRAND} 0%, {BRAND_DARK} 100%); }}
    .unduh-uat {{ background: linear-gradient(135deg, {NAVY_2} 0%, {NAVY} 100%); }}
    .unduh-kartu .ikon {{ font-size: 1.8rem; line-height: 1; }}
    .unduh-kartu .judul {{ font-size: 1.15rem; font-weight: 800; margin-top: 6px; }}
    .unduh-kartu .ket {{ font-size: 0.86rem; opacity: 0.92; margin-top: 2px; }}
    .unduh-kartu .isi {{ margin-top: 8px; }}
    .unduh-kartu .isi span {{
        display:inline-block; padding: 2px 9px; margin: 0 4px 4px 0; border-radius: 999px;
        background: rgba(255,255,255,0.18); border: 1px solid rgba(255,255,255,0.28);
        font-size: 0.74rem; font-weight: 800;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


KEY_TANGGAL = "tanggal_pengujian"
NAMA_HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]


def format_tanggal_singkat(tgl):
    """mis. date(2026, 11, 9) -> 'Senin, 9 November 2026'"""
    return f"{NAMA_HARI[tgl.weekday()]}, {tgl.day} {lampiran5.NAMA_BULAN[tgl.month - 1].capitalize()} {tgl.year}"


def stat_card(label, value, kind):
    st.markdown(f'<div class="stat stat-{kind}"><div class="lbl">{label}</div>'
                f'<div class="val">{value}</div></div>', unsafe_allow_html=True)


def chip(k, v, kosong_teks="belum ada"):
    nilai = v if v else kosong_teks
    kelas = "v" if v else "v kosong"
    return f'<div class="chip"><div class="k">{k}</div><div class="{kelas}">{nilai}</div></div>'


def judul_langkah(nomor, judul, sub):
    st.markdown(f'<div class="judul-langkah"><span class="no">{nomor}</span>{judul}</div>'
                f'<div class="sub-langkah">{sub}</div>', unsafe_allow_html=True)


# ============================================================
# SIDEBAR — PENGATURAN TETAP
# ============================================================

@st.cache_data(show_spinner=False)
def versi_aplikasi():
    """Versi kode (commit git) agar pengguna bisa memastikan sudah ter-update."""
    import subprocess
    try:
        out = subprocess.run(
            ["git", "log", "-1", "--format=%h · %cd", "--date=format:%d-%m-%Y %H:%M"],
            cwd=Path(__file__).resolve().parent, capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=5)
        return out.stdout.strip() if out.returncode == 0 else ""
    except Exception:  # noqa: BLE001
        return ""


_peng = muat_pengaturan()
with st.sidebar:
    st.markdown(
        f'<div class="sb-logo-card">{logo_img(52)}</div>'
        '<div class="sb-app">Automation Dokumen ASPI</div>'
        '<div class="sb-app-sub">UAT Script → Lampiran 5A–7C & UAT Result</div>',
        unsafe_allow_html=True,
    )
    if versi_aplikasi():
        st.caption(f"Versi: {versi_aplikasi()}")
    st.divider()
    sidebar_mitra = st.container()      # diisi setelah UAT Script diunggah
    st.divider()

    st.markdown('<div class="sb-section">⚙️ Pengaturan tetap</div>', unsafe_allow_html=True)
    st.caption("Jarang berubah. Dipakai di Lampiran 5A, 6A & 7A (penandatangan) "
               "dan 5B (user DevSite bank).")
    jabatan_1 = st.text_input("Jabatan penandatangan 1", value=_peng["jabatan_1"])
    nama_1 = st.text_input("Nama penandatangan 1", value=_peng["nama_1"])
    jabatan_2 = st.text_input("Jabatan penandatangan 2", value=_peng["jabatan_2"])
    nama_2 = st.text_input("Nama penandatangan 2", value=_peng["nama_2"])
    user_devsite_bank = st.text_input("User DevSite bank (Lampiran 5B)", value=_peng["user_devsite"])
    if st.button("💾 Simpan pengaturan", use_container_width=True):
        simpan_pengaturan({"jabatan_1": jabatan_1, "nama_1": nama_1,
                           "jabatan_2": jabatan_2, "nama_2": nama_2,
                           "user_devsite": user_devsite_bank})
        st.success("Tersimpan untuk penggunaan berikutnya.")

    st.divider()
    with st.expander("💡 Tips"):
        st.markdown(
            "- Nomor surat diambil dari **Nomor Referensi** UAT Script.\n"
            "- Tanggal pengujian 5A, 6A & 7A diatur di **langkah 2** (default hari ini).\n"
            "- UAT Result: di Word klik kanan Daftar Isi → **Update Field**.\n"
            "- Aplikasi berjalan **100% lokal**, data tidak dikirim ke internet."
        )

penandatangan = {"jabatan_1": jabatan_1, "nama_1": nama_1,
                 "jabatan_2": jabatan_2, "nama_2": nama_2}


# ============================================================
# HERO
# ============================================================

st.markdown(
    f"""
    <div class="hero">
        <div class="hero-logo-card">{logo_img(46)}</div>
        <div class="hero-teks">
            <h1>Automation Dokumen ASPI</h1>
            <p>Unggah UAT Script, cek API yang digunakan, lalu unduh dokumen dalam hitungan detik.</p>
            <span class="hero-tag">📚 Lampiran 5A–7C · 1 klik</span>
            <span class="hero-tag">📘 UAT Result · terpisah</span>
            <span class="hero-tag">🔒 100% lokal</span>
        </div>
        <div class="hero-kanan"><b>{date.today().day:02d}</b>{lampiran5.NAMA_BULAN[date.today().month - 1].capitalize()} {date.today().year}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

stepper_slot = st.empty()   # diisi setelah status tiap langkah diketahui


# ============================================================
# LANGKAH 1 — UNGGAH
# ============================================================

with st.container(border=True):
    judul_langkah(1, "Unggah UAT Script", "File Excel hasil UAT (.xlsx). Informasi mitra dibaca otomatis.")
    c_up, c_prod = st.columns([3, 2])
    with c_up:
        uploaded_file = st.file_uploader("File UAT Script (.xlsx)", type=["xlsx"],
                                         label_visibility="collapsed")
    with c_prod:
        product_label = st.selectbox(
            "Jenis produk", options=list(PRODUCT_OPTIONS.keys()), index=0,
            help="Biarkan 'Deteksi otomatis', atau paksa VA/FT / QRIS bila deteksi keliru.",
        )
    selected_product = PRODUCT_OPTIONS[product_label]

    analisa = None
    sig = None
    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        sig = f"{uploaded_file.name}|{len(file_bytes)}|{selected_product}"
        analisa = analisa_uat(file_bytes)
        produk = selected_product if selected_product != "AUTO" else analisa["produk"]
        if not analisa["ok"]:
            st.error("File tidak dapat dibaca sebagai UAT Script. Pastikan sheet **UAT Script** ada.")
            with st.expander("Detail teknis"):
                st.code(analisa["error"])
        else:
            st.markdown(
                '<div class="chips">'
                + chip("Mitra", analisa["pengguna"])
                + chip("Nomor referensi", analisa["nomor"])
                + chip("Produk", PRODUCT_LABEL.get(produk, produk or "-"))
                + chip("Tanggal pengujian", format_tanggal_singkat(st.session_state.get(KEY_TANGGAL, date.today())))
                + "</div>",
                unsafe_allow_html=True,
            )
            st.caption(f"📊 {analisa['jumlah_section']} section · {analisa['jumlah_case']} case terbaca "
                       f"dari **{uploaded_file.name}**")
    else:
        produk = None

file_ok = analisa is not None and analisa["ok"]
lampiran_5_6 = file_ok and produk in lampiran5.PRODUK_DIDUKUNG

# --- Sidebar: data mitra (berbeda tiap mitra) ---
user_devsite_mitra = ""
with sidebar_mitra:
    st.markdown('<div class="sb-section">🧾 Data mitra</div>', unsafe_allow_html=True)
    if not file_ok:
        st.caption("Unggah UAT Script untuk mengisi data mitra.")
    else:
        st.markdown(
            '<div class="sb-box"><div class="k">Mitra</div>'
            f'<div class="v">{analisa["pengguna"] or "-"}</div></div>',
            unsafe_allow_html=True,
        )
        if lampiran_5_6:
            user_devsite_mitra = st.text_input(
                "User pada Devsite mitra (Lampiran 6B) *", value="", key=f"user6_{sig}",
                placeholder="nama@mitra.co.id",
                help="Email user DevSite milik mitra; berbeda tiap mitra.",
            )
            if not user_devsite_mitra.strip():
                st.caption("⚠️ Wajib diisi agar kolom User pada Devsite di 6B tidak kosong.")
        else:
            st.caption("Produk QRIS: Lampiran 6B belum tersedia.")


# ============================================================
# LANGKAH 2 — VALIDASI API & DATA MITRA
# ============================================================

sub_api_tampil = None
tanda = None
pengguna = ""
nomor_surat = ""
if KEY_TANGGAL not in st.session_state:
    st.session_state[KEY_TANGGAL] = date.today()
tanggal_uji = st.session_state[KEY_TANGGAL]

with st.container(border=True):
    judul_langkah(2, "Validasi API & data mitra",
                  "Sub API yang digunakan dideteksi otomatis dari UAT Script, untuk Lampiran 5B, 6B & 7B.")
    if not file_ok:
        st.info("⬆️ Unggah UAT Script terlebih dahulu.")
    elif not lampiran_5_6:
        st.info("ℹ️ Produk **QRIS**: Lampiran 5A/5B, 6A/6B & 7A/7B belum tersedia (daftar Sub API QRIS belum ada). "
                "Lampiran 7C & UAT Result tetap dibuat.")
    else:
        validasi = analisa["validasi"]
        n_ok = sum(v["status"] == lampiran5.STATUS_DIGUNAKAN for v in validasi)
        n_no = sum(v["status"] == lampiran5.STATUS_TIDAK for v in validasi)
        n_cek = sum(v["status"] == lampiran5.STATUS_CEK for v in validasi)
        st.markdown(
            f'<span class="pill pill-ok">✅ {n_ok} digunakan</span>'
            f'<span class="pill pill-no">⛔ {n_no} tidak digunakan</span>'
            + (f'<span class="pill pill-cek">⚠️ {n_cek} perlu dicek</span>' if n_cek else ""),
            unsafe_allow_html=True,
        )
        if n_cek:
            st.warning("Ada Sub API yang **Nama Modul**-nya tidak dikenali di UAT Script. "
                       "Statusnya diperkirakan — mohon cek kolom *Lampirkan*.")

        tabel = st.data_editor(
            [{
                "Lampirkan": v["digunakan"],
                "Status": STATUS_LABEL[v["status"]],
                "Layanan API": v["layanan_api"],
                "Sub API": v["sub_api"],
                "Bukti dari UAT": v["bukti"],
                "Positif Req": True, "Positif Resp": True,
                "Negatif Req": True, "Negatif Resp": True,
            } for v in validasi],
            column_config={
                "Lampirkan": st.column_config.CheckboxColumn(
                    "Lampirkan", help="Centang = Sub API dimasukkan ke Lampiran 5B, 6B & 7B", width="small"),
                "Status": st.column_config.TextColumn("Status UAT", width="medium"),
                "Bukti dari UAT": st.column_config.TextColumn("Bukti (no. case & status)", width="large"),
                "Positif Req": st.column_config.CheckboxColumn("Pos Req", help="Centang = Ѵ, kosong = X", width="small"),
                "Positif Resp": st.column_config.CheckboxColumn("Pos Resp", help="Centang = Ѵ, kosong = X", width="small"),
                "Negatif Req": st.column_config.CheckboxColumn("Neg Req", help="Centang = Ѵ, kosong = X", width="small"),
                "Negatif Resp": st.column_config.CheckboxColumn("Neg Resp", help="Centang = Ѵ, kosong = X", width="small"),
            },
            disabled=["Status", "Layanan API", "Sub API", "Bukti dari UAT"],
            hide_index=True, use_container_width=True, key=f"validasi_{sig}",
        )
        st.caption("Kolom **Pos/Neg Req/Resp**: centang = **Ѵ** (sudah sesuai), kosong = **X** (belum sesuai). "
                   "Isian ini dipakai di Lampiran 5B, 6B & 7B.")

        sub_api_tampil = [r["Sub API"] for r in tabel if r["Lampirkan"]]
        tanda = [[bool(r["Positif Req"]), bool(r["Positif Resp"]),
                  bool(r["Negatif Req"]), bool(r["Negatif Resp"])] for r in tabel]
        if not sub_api_tampil:
            st.warning("Belum ada Sub API yang dicentang — tabel 5B & 6B akan menampilkan **semua** Sub API.")
            sub_api_tampil = None

    if file_ok:
        st.markdown("**Data dokumen**")
        if lampiran_5_6:
            m1, m2, m3 = st.columns([3, 2, 2])
        else:
            m1, m3 = st.columns([3, 2])
            m2 = None
        with m1:
            pengguna = st.text_input(
                "Pengguna Layanan (6B, 7B & 7C)" if lampiran_5_6 else "Nama Pengguna Layanan (7C)",
                value=analisa["pengguna"], key=f"pengguna_{sig}",
                help="Otomatis dari judul UAT Script; bisa diedit.")
        if m2 is not None:
            with m2:
                nomor_surat = st.text_input("Nomor surat (5A/5B)", value=analisa["nomor"],
                                            key=f"nomor_{sig}", help="Otomatis dari Nomor Referensi UAT Script.")
        with m3:
            tanggal_uji = st.date_input(
                "Tanggal pengujian (5A, 6A, 7A & 7C)" if lampiran_5_6 else "Tanggal Pengujian (7C)",
                key=KEY_TANGGAL, format="DD/MM/YYYY",
                help="Satu tanggal untuk semua lampiran. Nama hari dihitung otomatis.",
            )
        pratinjau_7c = (f'📄 Di 7C tertulis: <i>Nama Pengguna Layanan : '
                        f'{html.escape((pengguna or "").strip() or "-")} · Tanggal Pengujian : '
                        f'{format_tanggal_7c(tanggal_uji)}</i>')
        if lampiran_5_6:
            st.markdown(
                f'<div class="tgl-pratinjau">📅 Di 5A, 6A &amp; 7A tertulis: <i>"Pada hari '
                f'{lampiran5.format_tanggal_pengujian(tanggal_uji)}, telah selesai dilakukan …"</i>'
                f'<br>{pratinjau_7c}</div>',
                unsafe_allow_html=True,
            )
            st.caption("👈 **User Devsite mitra** untuk Lampiran 6B diisi di panel kiri.")
        else:
            st.markdown(f'<div class="tgl-pratinjau">{pratinjau_7c}</div>', unsafe_allow_html=True)


# ============================================================
# LANGKAH 3 — BUAT & UNDUH
# ============================================================

if "hasil" not in st.session_state:
    st.session_state.hasil = None

with st.container(border=True):
    judul_langkah(3, "Buat & unduh dokumen",
                  "Satu klik unduh semua Lampiran 5A–7C (file terpisah, .zip); UAT Result terpisah.")

    if file_ok:
        siap = [("✅", "UAT Script terbaca")]
        siap.append(("✅" if (pengguna or "").strip() else "⚠️",
                     "Pengguna Layanan" + ("" if (pengguna or "").strip() else " belum ada")))
        if not lampiran_5_6:
            siap.append(("📅", f"Tanggal 7C: {format_tanggal_singkat(tanggal_uji)}"))
        if lampiran_5_6:
            siap.append(("✅" if nomor_surat.strip() else "⚠️",
                         "Nomor surat" + ("" if nomor_surat.strip() else " belum ada")))
            siap.append(("📅", f"Tanggal 5A/6A/7A/7C: {format_tanggal_singkat(tanggal_uji)}"))
            siap.append(("✅" if user_devsite_mitra.strip() else "⚠️",
                         "User Devsite mitra" + ("" if user_devsite_mitra.strip()
                                                 else " belum diisi — isi di panel kiri")))
        st.markdown(" &nbsp;·&nbsp; ".join(f"{a} {b}" for a, b in siap), unsafe_allow_html=True)

    buat = st.button("🚀 Buat semua dokumen", type="primary", disabled=not file_ok,
                     use_container_width=True)

    if buat and file_ok:
        try:
            with st.spinner("Membuat dokumen..."):
                (lampiran_bytes, stats, warnings, product_key,
                 lampiran_name, status_counts) = convert_uat_to_lampiran_bytes(
                    file_bytes, product=selected_product,
                    pengguna_layanan=(pengguna or "").strip() or analisa["pengguna"],
                    tanggal_pengujian=tanggal_uji)
                result_bytes, _pk, result_name = convert_uat_to_result_bytes(
                    file_bytes, product=selected_product)
                lamp5 = lampiran5.convert_lampiran_5_bytes(
                    file_bytes, product_key, penandatangan, user_devsite_bank,
                    tanda=tanda, nomor_surat=nomor_surat, sub_api_tampil=sub_api_tampil,
                    tanggal=tanggal_uji)
                lamp6 = lampiran6.convert_lampiran_6_bytes(
                    file_bytes, product_key, penandatangan, pengguna, user_devsite_mitra,
                    tanda=tanda, sub_api_tampil=sub_api_tampil, tanggal=tanggal_uji)
                lamp7 = lampiran7.convert_lampiran_7_bytes(
                    file_bytes, product_key, penandatangan, pengguna,
                    tanda=tanda, sub_api_tampil=sub_api_tampil, tanggal=tanggal_uji)
                nama_pt = (pengguna or "").strip() or (analisa["pengguna"] if analisa else "")
            st.session_state.hasil = {
                "sig": sig, "produk": product_key, "stats": stats, "warnings": warnings,
                "status": status_counts, "lamp5": lamp5, "lamp6": lamp6, "lamp7": lamp7,
                "dok_utama": [(lampiran_name, lampiran_bytes), (result_name, result_bytes)],
                "nama_pt": nama_pt,
                "tanggal": tanggal_uji,
                "waktu": datetime.now().strftime("%Y%m%d_%H%M%S"),
            }
            st.toast("Dokumen berhasil dibuat!", icon="🎉")
        except Exception as e:  # noqa: BLE001
            st.session_state.hasil = None
            st.error("Gagal membuat dokumen. Pastikan file adalah UAT Script (.xlsx) yang benar.")
            with st.expander("Detail teknis"):
                st.code(str(e))

    hasil = st.session_state.hasil
    hasil_berlaku = (hasil is not None and hasil["sig"] == sig
                     and hasil.get("tanggal") == tanggal_uji)
    if hasil is not None and not hasil_berlaku and file_ok:
        st.info("File, produk, atau tanggal pengujian berubah — klik **Buat semua dokumen** lagi.")

    if hasil_berlaku:
        ts = hasil["waktu"]

        def beri_waktu(nama):
            return nama.replace(".docx", f" {ts}.docx")

        (l7_name, l7_bytes), (ur_name, ur_bytes) = hasil["dok_utama"]

        # Daftar file Lampiran (urut 5A ... 7C); UAT Result terpisah
        lampiran = []
        if hasil["lamp5"]:
            l5 = hasil["lamp5"]
            lampiran += [
                ("📝", "5A", f"Informasi uji Developer Site · No. {l5['nomor_surat'] or '-'}",
                 beri_waktu(l5["a_name"]), l5["a_bytes"]),
                ("🗂️", "5B", f"Matriks asesmen Developer Site · {len(l5['sub_api'])} Sub API",
                 beri_waktu(l5["b_name"]), l5["b_bytes"]),
            ]
        if hasil["lamp6"]:
            l6 = hasil["lamp6"]
            lampiran += [
                ("📝", "6A", "Informasi uji Developer Site Pengguna Layanan.",
                 beri_waktu(l6["a_name"]), l6["a_bytes"]),
                ("🗂️", "6B", f"Matriks asesmen {l6['pengguna'] or 'Pengguna Layanan'}",
                 beri_waktu(l6["b_name"]), l6["b_bytes"]),
            ]
        if hasil.get("lamp7"):
            l7 = hasil["lamp7"]
            lampiran += [
                ("🧾", "7A", "Berita acara uji fungsionalitas.",
                 beri_waktu(l7["a_name"]), l7["a_bytes"]),
                ("✅", "7B", f"Matriks uji fungsionalitas · {len(l7['sub_api'])} Sub API",
                 beri_waktu(l7["b_name"]), l7["b_bytes"]),
            ]
        lampiran.append(("📊", "7C", "Tabel skenario & hasil uji fungsionalitas.",
                         beri_waktu(l7_name or OUTPUT_FILENAME), l7_bytes))

        # Satu tombol: semua file Lampiran (masing-masing file Word terpisah) dalam 1 .zip
        kode = [x[1] for x in lampiran]
        rentang = kode[0] if len(kode) == 1 else f"{kode[0]}-{kode[-1]}"
        nama_pt = hasil.get("nama_pt") or ""
        zip_nama = (f"Lampiran ASPI {rentang}"
                    + (f" - {lampiran5.nama_file_aman(nama_pt)}" if nama_pt else "") + f" {ts}.zip")
        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
            for _ikon, _kode, _ket, nama, data in lampiran:
                zf.writestr(nama, data)

        st.success("✅ Dokumen siap diunduh.")
        u1, u2 = st.columns([3, 2])
        with u1:
            st.markdown(
                '<div class="unduh-kartu unduh-lampiran"><div class="ikon">📚</div>'
                '<div class="judul">Semua Lampiran ASPI</div>'
                f'<div class="ket">{len(lampiran)} file Word terpisah · sekali klik (.zip)</div>'
                '<div class="isi">' + "".join(f"<span>{k}</span>" for k in kode) + "</div></div>",
                unsafe_allow_html=True,
            )
            st.download_button(f"⬇️ Unduh semua Lampiran ({len(lampiran)} file, .zip)", data=zbuf.getvalue(),
                               file_name=zip_nama, mime="application/zip", type="primary",
                               use_container_width=True, key="dl_semua_lampiran")
        with u2:
            st.markdown(
                '<div class="unduh-kartu unduh-uat"><div class="ikon">📘</div>'
                '<div class="judul">UAT Result</div>'
                '<div class="ket">Dokumen terpisah · semua case + Daftar Isi · logo BSS di header</div>'
                '<div class="isi"><span>UAT Result</span></div></div>',
                unsafe_allow_html=True,
            )
            st.download_button("⬇️ Unduh UAT Result (.docx)", data=ur_bytes,
                               file_name=beri_waktu(ur_name or "UAT Result.docx"), mime=MIME_DOCX,
                               use_container_width=True, key="dl_uat_result")

        with st.expander("📂 Unduh per lampiran (opsional)"):
            for i in range(0, len(lampiran), 4):
                cols = st.columns(4)
                for col, (ikon, kd, ket, nama, data) in zip(cols, lampiran[i:i + 4]):
                    with col, st.container(border=True):
                        st.markdown(f'<div class="doc-ikon">{ikon}</div><div class="doc-judul">Lampiran {kd}</div>'
                                    f'<div class="doc-ket">{ket}</div>', unsafe_allow_html=True)
                        st.download_button("⬇️ Unduh .docx", data=data, file_name=nama, mime=MIME_DOCX,
                                           use_container_width=True, key=f"dl_{kd}")

        if hasil["lamp5"] and not hasil["lamp5"]["nomor_surat"]:
            st.warning("⚠️ Nomor surat kosong di 5A/5B — isi di langkah 2 lalu buat ulang.")
        if hasil["lamp6"] and not user_devsite_mitra.strip():
            st.warning("⚠️ User Devsite mitra kosong — isi di panel kiri, lalu buat ulang.")

        # ---------- Ringkasan hasil Lampiran 7C ----------
        status = hasil["status"] or {}
        stats = hasil["stats"] or {}
        warns = hasil["warnings"] or []
        tot_pass = sum(d.get("PASS", 0) for d in status.values())
        tot_fail = sum(d.get("NOT PASS", 0) for d in status.values())
        tot_na = sum(d.get("N/A", 0) for d in status.values())
        tot_all = tot_pass + tot_fail + tot_na

        st.markdown("**Ringkasan hasil uji (Lampiran 7C)**")
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            stat_card("Total baris", tot_all, "tot")
        with k2:
            stat_card("PASS", tot_pass, "pass")
        with k3:
            stat_card("NOT PASS", tot_fail, "fail")
        with k4:
            stat_card("N/A", tot_na, "na")
        st.write("")

        tab_cek, tab_grafik, tab_ringkas = st.tabs(
            [f"⚠️ Perlu dicek ({len(warns)})", "📊 Grafik", "📋 Per layanan"])
        with tab_cek:
            if warns:
                st.caption("Data TIDAK diubah — parameter & value tetap apa adanya. Mohon cek manual:")
                for w in warns:
                    st.markdown(f"- {w}")
            else:
                st.success("Tidak ada data abnormal terdeteksi.")
        with tab_grafik:
            aktif = {s: d for s, d in status.items() if sum(d.values()) > 0}
            if aktif:
                try:
                    import pandas as pd
                    df = pd.DataFrame({k: [aktif[s][k] for s in aktif] for k in ("PASS", "NOT PASS", "N/A")},
                                      index=list(aktif.keys()))
                    st.bar_chart(df, color=["#2e9e4f", "#e24b4b", "#f0a93b"],
                                 use_container_width=True, height=320)
                except Exception:  # noqa: BLE001
                    st.bar_chart({k: [aktif[s][k] for s in aktif] for k in ("PASS", "NOT PASS", "N/A")},
                                 use_container_width=True, height=320)
                if tot_all:
                    st.caption(f"PASS {tot_pass/tot_all*100:.0f}% · NOT PASS {tot_fail/tot_all*100:.0f}% · "
                               f"N/A {tot_na/tot_all*100:.0f}%")
            else:
                st.caption("Tidak ada data untuk ditampilkan.")
        with tab_ringkas:
            baris = []
            for section_name, _total in PRODUCT_PROFILES[hasil["produk"]]["sections"]:
                sc = status.get(section_name, {})
                terisi = stats.get(section_name, 0)
                baris.append({"Layanan API": section_name, "PASS": sc.get("PASS", 0),
                              "NOT PASS": sc.get("NOT PASS", 0), "N/A": sc.get("N/A", 0),
                              "Total baris": terisi,
                              "Status": "✅ Ditampilkan" if terisi > 0 else "— Tidak dites"})
            st.dataframe(baris, use_container_width=True, hide_index=True)


# ============================================================
# STEPPER (status tiap langkah)
# ============================================================

def _step(n, teks, kelas):
    tanda = "✓" if kelas == "done" else str(n)
    return f'<div class="step {kelas}"><div class="num">{tanda}</div>{teks}</div>'


_hasil_ok = st.session_state.hasil is not None and st.session_state.hasil.get("sig") == sig
stepper_slot.markdown(
    '<div class="stepper">'
    + _step(1, "Unggah UAT Script", "done" if file_ok else "active")
    + _step(2, "Validasi API & data mitra", "done" if _hasil_ok else ("active" if file_ok else ""))
    + _step(3, "Unduh dokumen", "done" if _hasil_ok else "")
    + "</div>",
    unsafe_allow_html=True,
)


# ============================================================
# FOOTER
# ============================================================

with st.expander("ℹ️ Tentang aplikasi ini"):
    st.markdown(
        """
        **Dokumen yang dibuat** — semua Lampiran 5A–7C diunduh **sekali klik** (.zip berisi
        file Word terpisah per lampiran); **UAT Result** diunduh terpisah (logo BSS di header).
        - **Lampiran 7C** — tabel skenario hasil uji fungsional (Tidak dites → `N/A`, Remark ke Notes).
        - **UAT Result** — semua case & section + Daftar Isi otomatis.
        - **Lampiran 5A/5B** (VA/FT) — Developer Site Penyedia Layanan; nomor surat dari Nomor Referensi UAT.
        - **Lampiran 6A/6B** (VA/FT) — Developer Site Pengguna Layanan; nama PT dari judul UAT.
        - **Lampiran 7A/7B** (VA/FT) — Berita acara & matriks uji fungsionalitas dengan Pengguna Layanan.

        **Validasi API**: Sub API dianggap *digunakan* bila ada case dengan Nama Modul yang sesuai
        berstatus selain *Tidak dites*. Nomor case & status ditampilkan sebagai bukti.
        """
    )
    st.caption("© Automation Dokumen ASPI — Bank Sahabat Sampoerna")

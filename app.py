"""
Aplikasi Web Lampiran 7C Automation
===================================
Aplikasi web berbasis Streamlit untuk mengubah file UAT Script (Excel) menjadi
dokumen Lampiran 7C ASPI dan UAT Result (Word).

Cara pakai:
    1. Install dependency:  pip install -r requirements.txt
    2. Jalankan aplikasi:   streamlit run app.py
    3. Browser akan terbuka otomatis. Pilih produk, unggah file UAT Script,
       klik "Proses", lalu unduh kedua dokumen hasilnya.
"""

from datetime import datetime

import streamlit as st

from main import (
    convert_uat_to_lampiran_bytes,
    convert_uat_to_result_bytes,
    PRODUCT_PROFILES,
    OUTPUT_FILENAME,
)

# Label produk untuk dropdown (tampilan -> kunci di PRODUCT_PROFILES / "AUTO")
PRODUCT_OPTIONS = {
    "Deteksi otomatis": "AUTO",
    "Fund Transfer & Virtual Account (VA/FT)": "FT_VA",
    "QRIS": "QRIS",
}
PRODUCT_LABEL = {"FT_VA": "Fund Transfer & Virtual Account (VA/FT)", "QRIS": "QRIS"}

# Palet warna brand
BRAND = "#ED7D31"       # oranye ASPI
BRAND_DARK = "#C55A11"
OK_COLOR = "#2e7d32"
WARN_COLOR = "#b26a00"


# ============================================================
# KONFIGURASI HALAMAN
# ============================================================

st.set_page_config(
    page_title="Automation Lampiran 7C ASPI",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# CSS KUSTOM (tema korporat, kartu, bayangan halus)
# ============================================================

st.markdown(
    f"""
    <style>
    .block-container {{ padding-top: 1.6rem; max-width: 1180px; }}

    /* Hero header */
    .hero {{
        background: linear-gradient(120deg, {BRAND} 0%, {BRAND_DARK} 100%);
        padding: 30px 38px; border-radius: 18px; color: #fff;
        box-shadow: 0 12px 30px rgba(237,125,49,0.30);
        display: flex; align-items: center; gap: 22px;
    }}
    .hero-logo {{
        width: 64px; height: 64px; border-radius: 16px;
        background: rgba(255,255,255,0.18);
        display:flex; align-items:center; justify-content:center;
        font-size: 34px; flex-shrink: 0;
        border: 1px solid rgba(255,255,255,0.35);
    }}
    .hero h1 {{ color:#fff; font-size: 1.85rem; margin: 0 0 4px 0; font-weight: 800; letter-spacing:-0.3px; }}
    .hero p  {{ color:#fff8f2; font-size: 1.0rem; margin: 0; opacity: 0.96; }}

    /* Kartu umum */
    .card {{
        background: var(--secondary-background-color, #ffffff);
        border: 1px solid rgba(0,0,0,0.07);
        border-radius: 14px; padding: 18px 20px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
        margin-bottom: 4px;
    }}

    /* Kartu langkah */
    .step-card {{
        background: var(--secondary-background-color, #faf8f6);
        border: 1px solid rgba(0,0,0,0.06);
        border-left: 5px solid {BRAND};
        border-radius: 12px; padding: 13px 18px; margin: 10px 0 2px 0;
    }}
    .step-num {{
        display:inline-flex; align-items:center; justify-content:center;
        width: 26px; height: 26px; border-radius: 50%;
        background:{BRAND}; color:#fff; font-weight:700; font-size:0.85rem; margin-right: 10px;
    }}
    .step-title {{ font-weight: 700; font-size: 1.08rem; }}

    /* Badge */
    .badge {{ display:inline-block; padding: 3px 13px; border-radius: 999px; font-size: 0.78rem; font-weight: 700; }}
    .badge-ftva {{ background:#e8f0fe; color:#1a56db; }}
    .badge-qris {{ background:#e9f9ee; color:#177245; }}
    .badge-auto {{ background:#fdf0e6; color:{BRAND_DARK}; }}

    /* Stat card */
    .stat {{
        border-radius: 14px; padding: 16px 18px; color:#fff; text-align:left;
        box-shadow: 0 4px 14px rgba(0,0,0,0.10);
    }}
    .stat .lbl {{ font-size: 0.82rem; opacity: 0.92; font-weight:600; margin-bottom:2px; }}
    .stat .val {{ font-size: 1.9rem; font-weight: 800; line-height:1; }}
    .stat-pass {{ background: linear-gradient(135deg,#2e9e4f,#1e7e38); }}
    .stat-fail {{ background: linear-gradient(135deg,#e24b4b,#c62828); }}
    .stat-na   {{ background: linear-gradient(135deg,#f0a93b,#d9871f); }}
    .stat-tot  {{ background: linear-gradient(135deg,#5b6b7a,#3c4753); }}

    /* Tombol */
    .stButton>button, .stDownloadButton>button {{
        border-radius: 11px; font-weight: 700; padding: 0.6rem 1rem;
    }}
    div[data-testid="stMetric"] {{
        background: var(--secondary-background-color, #faf8f6);
        border: 1px solid rgba(0,0,0,0.06); border-radius: 12px; padding: 12px 16px;
    }}
    /* File uploader lebih menonjol */
    section[data-testid="stFileUploaderDropzone"] {{
        border: 2px dashed {BRAND}; border-radius: 14px; background: #fff8f2;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


def stat_card(label, value, kind):
    st.markdown(
        f'<div class="stat stat-{kind}"><div class="lbl">{label}</div>'
        f'<div class="val">{value}</div></div>',
        unsafe_allow_html=True,
    )


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.markdown(f"<h3 style='margin-bottom:0'>📄 Lampiran 7C ASPI</h3>", unsafe_allow_html=True)
    st.caption("Automation UAT Script → Lampiran 7C & UAT Result")
    st.divider()

    st.markdown("#### 🧭 Langkah singkat")
    st.markdown(
        "1. Pilih **jenis produk**\n"
        "2. **Unggah** file UAT Script (.xlsx)\n"
        "3. Klik **Proses**\n"
        "4. **Unduh** 2 dokumen hasil"
    )
    st.divider()

    st.markdown("#### 📦 Produk didukung")
    st.markdown(
        "- **VA/FT** — Balance, Intrabank, Interbank, RTGS, SKNBI, Virtual Account\n"
        "- **QRIS** — Balance Inquiry, Transaction History, QR MPM"
    )
    st.divider()
    st.caption("💡 UAT Result: buka di Word → klik kanan Daftar Isi → *Update Field* "
               "untuk memunculkan nomor halaman.")
    st.caption("🔒 Berjalan 100% lokal — data tidak dikirim ke internet.")


# ============================================================
# HERO HEADER
# ============================================================

st.markdown(
    """
    <div class="hero">
        <div class="hero-logo">📄</div>
        <div>
            <h1>Automation Lampiran 7C ASPI</h1>
            <p>Ubah UAT Script (Excel) menjadi Lampiran 7C &amp; UAT Result (Word) secara otomatis —
            unggah, proses, unduh. Mendukung Fund Transfer/VA dan QRIS.</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.write("")


# ============================================================
# LANGKAH 1 — PILIH PRODUK & UNGGAH
# ============================================================

st.markdown(
    '<div class="step-card"><span class="step-num">1</span>'
    '<span class="step-title">Pilih produk &amp; unggah file UAT Script</span></div>',
    unsafe_allow_html=True,
)

col_a, col_b = st.columns([1, 1])

with col_a:
    product_label = st.selectbox(
        "Jenis produk UAT Script",
        options=list(PRODUCT_OPTIONS.keys()),
        index=0,
        help=(
            "Pilih 'Fund Transfer & Virtual Account' untuk UAT VA/FT, atau 'QRIS' "
            "untuk UAT QRIS. Pilihan ini MEMAKSA konfigurasi produk. Gunakan "
            "'Deteksi otomatis' bila ingin aplikasi menebak sendiri dari isi file."
        ),
    )
    selected_product = PRODUCT_OPTIONS[product_label]

    if selected_product == "FT_VA":
        st.markdown('<span class="badge badge-ftva">MODE VA/FT</span>', unsafe_allow_html=True)
        st.caption("Section: Balance Inquiry, Intrabank, Interbank, RTGS, SKNBI, Virtual Account.")
    elif selected_product == "QRIS":
        st.markdown('<span class="badge badge-qris">MODE QRIS</span>', unsafe_allow_html=True)
        st.caption("Section: Balance Inquiry, API Transaction History List, QR MPM.")
    else:
        st.markdown('<span class="badge badge-auto">DETEKSI OTOMATIS</span>', unsafe_allow_html=True)
        st.caption("Produk ditentukan dari isi file.")

with col_b:
    uploaded_file = st.file_uploader(
        "Pilih file UAT Script (.xlsx)",
        type=["xlsx"],
        help="File Excel hasil UAT. Harus memiliki sheet berisi data UAT Script.",
    )
    if uploaded_file is not None:
        st.success(f"✅ **{uploaded_file.name}**  \n{uploaded_file.size / 1024:.1f} KB")

st.info("📄 Sekali proses menghasilkan **2 dokumen**: **Lampiran 7C** dan **UAT Result**.")


# ============================================================
# LANGKAH 2 — PROSES
# ============================================================

st.markdown(
    '<div class="step-card"><span class="step-num">2</span>'
    '<span class="step-title">Proses &amp; buat dokumen</span></div>',
    unsafe_allow_html=True,
)

process_clicked = st.button(
    "🚀 Proses & Buat Lampiran 7C + UAT Result",
    type="primary",
    disabled=(uploaded_file is None),
    use_container_width=True,
)
if uploaded_file is None:
    st.caption("⬆️ Unggah file terlebih dahulu untuk mengaktifkan tombol proses.")

# Simpan hasil di session_state agar tombol unduh tetap muncul setelah rerun
if "lampiran_bytes" not in st.session_state:
    st.session_state.lampiran_bytes = None
    st.session_state.lampiran_name = None
    st.session_state.result_docx_bytes = None
    st.session_state.result_docx_name = None
    st.session_state.result_stats = None
    st.session_state.result_warnings = None
    st.session_state.result_product = None
    st.session_state.result_status = None

if process_clicked and uploaded_file is not None:
    try:
        progress = st.progress(0, text="Membaca file...")
        file_bytes = uploaded_file.getvalue()

        progress.progress(30, text="Membuat Lampiran 7C...")
        (lampiran_bytes, stats, warnings, product_key,
         lampiran_name, status_counts) = convert_uat_to_lampiran_bytes(
            file_bytes, product=selected_product
        )

        progress.progress(70, text="Membuat UAT Result...")
        result_bytes, _pk2, result_name = convert_uat_to_result_bytes(
            file_bytes, product=selected_product
        )
        progress.progress(100, text="Selesai!")
        progress.empty()

        st.session_state.lampiran_bytes = lampiran_bytes
        st.session_state.lampiran_name = lampiran_name
        st.session_state.result_docx_bytes = result_bytes
        st.session_state.result_docx_name = result_name
        st.session_state.result_stats = stats
        st.session_state.result_warnings = warnings
        st.session_state.result_product = product_key
        st.session_state.result_status = status_counts

        st.toast("Dokumen berhasil dibuat!", icon="🎉")

        total = sum(stats.values())
        if total == 0:
            st.warning(
                "File berhasil diproses, tetapi **tidak ada data hasil UAT** "
                "yang ditemukan. Periksa kembali isi file UAT Script Anda "
                "atau pastikan jenis produk yang dipilih sudah tepat."
            )
        else:
            st.success(
                f"Berhasil! Produk **{PRODUCT_LABEL.get(product_key, product_key)}** — "
                f"total **{total} baris** data pada Lampiran 7C. Hasil siap diunduh di bawah."
            )
    except Exception as e:  # noqa: BLE001
        for k in ("lampiran_bytes", "result_docx_bytes", "result_stats",
                  "result_warnings", "result_product", "result_status"):
            st.session_state[k] = None
        st.error(
            "Gagal memproses file. Pastikan file yang diunggah adalah "
            "UAT Script (.xlsx) dengan format yang benar."
        )
        with st.expander("Detail teknis (untuk troubleshooting)"):
            st.code(str(e))


# ============================================================
# LANGKAH 3 — RINGKASAN & UNDUH
# ============================================================

if st.session_state.lampiran_bytes is not None:
    st.markdown(
        '<div class="step-card"><span class="step-num">3</span>'
        '<span class="step-title">Ringkasan &amp; unduh hasil</span></div>',
        unsafe_allow_html=True,
    )

    used_product = st.session_state.get("result_product") or "FT_VA"
    stats = st.session_state.result_stats or {}
    status = st.session_state.result_status or {}
    warns = st.session_state.get("result_warnings") or []

    # Hitung total status
    tot_pass = sum(d.get("PASS", 0) for d in status.values())
    tot_fail = sum(d.get("NOT PASS", 0) for d in status.values())
    tot_na = sum(d.get("N/A", 0) for d in status.values())
    tot_all = tot_pass + tot_fail + tot_na

    # --- Kartu statistik berwarna ---
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        stat_card("Total baris", tot_all, "tot")
    with c2:
        stat_card("PASS", tot_pass, "pass")
    with c3:
        stat_card("NOT PASS", tot_fail, "fail")
    with c4:
        stat_card("N/A", tot_na, "na")

    st.write("")

    # --- Tab: Unduh | Grafik | Ringkasan | Perlu dicek ---
    tab_unduh, tab_grafik, tab_ringkasan, tab_warning = st.tabs(
        ["⬇️ Unduh", "📊 Grafik", "📋 Ringkasan section", f"⚠️ Perlu dicek ({len(warns)})"]
    )

    with tab_unduh:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        mime_docx = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        lamp_name = (st.session_state.get("lampiran_name") or OUTPUT_FILENAME).replace(
            ".docx", f" {timestamp}.docx"
        )
        res_name = (st.session_state.get("result_docx_name") or "UAT Result.docx").replace(
            ".docx", f" {timestamp}.docx"
        )
        st.write("Kedua dokumen berhasil dibuat. Silakan unduh:")
        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                "⬇️ Lampiran 7C (.docx)",
                data=st.session_state.lampiran_bytes,
                file_name=lamp_name, mime=mime_docx,
                type="primary", use_container_width=True,
            )
            st.caption("Tabel skenario hasil uji fungsional (laporan ASPI).")
        with dl2:
            st.download_button(
                "⬇️ UAT Result (.docx)",
                data=st.session_state.result_docx_bytes,
                file_name=res_name, mime=mime_docx,
                type="secondary", use_container_width=True,
            )
            st.caption("Dokumen naratif per-case + Daftar Isi.")

    with tab_grafik:
        # Chart batang per section: PASS / NOT PASS / N/A
        active = {s: d for s, d in status.items() if sum(d.values()) > 0}
        if active:
            st.caption("Jumlah hasil per layanan API")
            try:
                import pandas as pd
                df = pd.DataFrame(
                    {
                        "PASS": [active[s]["PASS"] for s in active],
                        "NOT PASS": [active[s]["NOT PASS"] for s in active],
                        "N/A": [active[s]["N/A"] for s in active],
                    },
                    index=list(active.keys()),
                )
                st.bar_chart(
                    df,
                    color=["#2e9e4f", "#e24b4b", "#f0a93b"],  # urut kolom: PASS, NOT PASS, N/A
                    use_container_width=True,
                    height=340,
                )
            except Exception:
                # Fallback sederhana bila pandas/warna tidak tersedia
                st.bar_chart(
                    {
                        "PASS": [active[s]["PASS"] for s in active],
                        "NOT PASS": [active[s]["NOT PASS"] for s in active],
                        "N/A": [active[s]["N/A"] for s in active],
                    },
                    use_container_width=True, height=340,
                )
            # Proporsi keseluruhan
            if tot_all:
                st.caption(
                    f"Proporsi: PASS {tot_pass/tot_all*100:.0f}% · "
                    f"NOT PASS {tot_fail/tot_all*100:.0f}% · "
                    f"N/A {tot_na/tot_all*100:.0f}%"
                )
        else:
            st.caption("Tidak ada data untuk ditampilkan.")

    with tab_ringkasan:
        used_sections = PRODUCT_PROFILES[used_product]["sections"]
        rows = []
        for section_name, _total in used_sections:
            sc = status.get(section_name, {})
            terisi = stats.get(section_name, 0)
            rows.append({
                "Layanan API": section_name,
                "PASS": sc.get("PASS", 0),
                "NOT PASS": sc.get("NOT PASS", 0),
                "N/A": sc.get("N/A", 0),
                "Total baris": terisi,
                "Status": "✅ Ditampilkan" if terisi > 0 else "— Tidak dites",
            })
        st.dataframe(rows, use_container_width=True, hide_index=True)

    with tab_warning:
        if warns:
            st.warning(
                f"Ditemukan **{len(warns)} hal yang perlu dicek**. "
                "Data TIDAK diubah — parameter & value tetap apa adanya."
            )
            for w in warns:
                st.markdown(f"- {w}")
        else:
            st.success("✅ Tidak ada data abnormal terdeteksi. Parameter & value aman.")


# ============================================================
# FOOTER
# ============================================================

st.divider()
with st.expander("ℹ️ Tentang aplikasi ini"):
    st.markdown(
        """
        Aplikasi ini membaca file **UAT Script** lalu menghasilkan **2 dokumen**:

        **Lampiran 7C** (tabel skenario hasil uji fungsional):
        - Hanya layanan yang **benar-benar dites** yang ditampilkan.
        - Kolom **Scenario** diambil dari kolom *Langkah Tes*.
        - Kolom **Request** & **Response** dipisah dan diformat (Response di-*compress*).
        - Baris **Tidak dites / Belum dites** → hasil `N/A`, Remark dipindah ke Notes.

        **UAT Result** (dokumen naratif per-case):
        - Menampilkan **semua case & semua section** + **Daftar Isi** otomatis.
        - Log di-*compress* (URL dibiarkan, Header & body dipadatkan).

        Mendukung produk **Fund Transfer / Virtual Account** dan **QRIS**
        (deteksi otomatis atau dipilih manual).
        """
    )
    st.caption("© Automation Lampiran 7C ASPI — Bank Sahabat Sampoerna")

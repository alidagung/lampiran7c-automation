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
# CSS KUSTOM (tema oranye ASPI, kartu, tombol)
# ============================================================

st.markdown(
    """
    <style>
    /* Lebar konten utama dibatasi agar nyaman dibaca */
    .block-container { padding-top: 2rem; max-width: 1100px; }

    /* Hero header */
    .hero {
        background: linear-gradient(120deg, #ED7D31 0%, #C55A11 100%);
        padding: 28px 34px;
        border-radius: 16px;
        color: #ffffff;
        box-shadow: 0 8px 24px rgba(237,125,49,0.28);
        margin-bottom: 8px;
    }
    .hero h1 { color:#fff; font-size: 1.9rem; margin: 0 0 6px 0; font-weight: 800; }
    .hero p  { color: #fff8f2; font-size: 1.02rem; margin: 0; opacity: 0.95; }

    /* Kartu langkah */
    .step-card {
        background: var(--secondary-background-color, #f7f7f9);
        border: 1px solid rgba(0,0,0,0.06);
        border-left: 5px solid #ED7D31;
        border-radius: 12px;
        padding: 14px 18px;
        margin-bottom: 6px;
    }
    .step-num {
        display:inline-flex; align-items:center; justify-content:center;
        width: 26px; height: 26px; border-radius: 50%;
        background:#ED7D31; color:#fff; font-weight:700; font-size:0.85rem;
        margin-right: 8px;
    }
    .step-title { font-weight: 700; font-size: 1.05rem; }

    /* Badge produk */
    .badge {
        display:inline-block; padding: 3px 12px; border-radius: 999px;
        font-size: 0.8rem; font-weight: 700; letter-spacing: .2px;
    }
    .badge-ftva { background:#e8f0fe; color:#1a56db; }
    .badge-qris { background:#e9f9ee; color:#177245; }
    .badge-auto { background:#fdf0e6; color:#C55A11; }

    /* Tombol lebih tebal */
    .stButton>button, .stDownloadButton>button {
        border-radius: 10px; font-weight: 700; padding: 0.55rem 1rem;
    }
    /* Metric card */
    div[data-testid="stMetric"] {
        background: var(--secondary-background-color, #f7f7f9);
        border: 1px solid rgba(0,0,0,0.06);
        border-radius: 12px; padding: 12px 16px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.markdown("### 📄 Lampiran 7C ASPI")
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


# ============================================================
# HERO HEADER
# ============================================================

st.markdown(
    """
    <div class="hero">
        <h1>📄 Automation Lampiran 7C ASPI</h1>
        <p>Ubah file UAT Script (Excel) menjadi dokumen Lampiran 7C dan UAT Result (Word)
        secara otomatis — cukup unggah, proses, lalu unduh.</p>
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
    '<span class="step-title">Pilih produk & unggah file UAT Script</span></div>',
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
            "untuk UAT QRIS. Pilihan ini MEMAKSA konfigurasi produk yang dipakai. "
            "Gunakan 'Deteksi otomatis' bila ingin aplikasi menebak sendiri dari isi file."
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

st.info("📄 Sekali proses akan menghasilkan **2 dokumen**: **Lampiran 7C** dan **UAT Result**.")


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

if process_clicked and uploaded_file is not None:
    try:
        with st.spinner("Memproses file, membuat 2 dokumen..."):
            file_bytes = uploaded_file.getvalue()
            # 1) Lampiran 7C
            lampiran_bytes, stats, warnings, product_key, lampiran_name = (
                convert_uat_to_lampiran_bytes(file_bytes, product=selected_product)
            )
            # 2) UAT Result
            result_bytes, _pk2, result_name = convert_uat_to_result_bytes(
                file_bytes, product=selected_product
            )

        st.session_state.lampiran_bytes = lampiran_bytes
        st.session_state.lampiran_name = lampiran_name
        st.session_state.result_docx_bytes = result_bytes
        st.session_state.result_docx_name = result_name
        st.session_state.result_stats = stats
        st.session_state.result_warnings = warnings
        st.session_state.result_product = product_key

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
                f"total **{total} baris** data pada Lampiran 7C. "
                "Kedua dokumen siap diunduh di bawah."
            )
    except Exception as e:  # noqa: BLE001 - tampilkan pesan error ramah pengguna
        st.session_state.lampiran_bytes = None
        st.session_state.result_docx_bytes = None
        st.session_state.result_stats = None
        st.session_state.result_warnings = None
        st.session_state.result_product = None
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

    # --- Metrik ringkas ---
    if st.session_state.result_stats is not None:
        stats = st.session_state.result_stats
        total = sum(stats.values())
        jumlah_layanan = sum(1 for v in stats.values() if v > 0)
        warns = st.session_state.get("result_warnings") or []

        m1, m2, m3 = st.columns(3)
        m1.metric("Total baris data", total)
        m2.metric("Layanan ditampilkan", jumlah_layanan)
        m3.metric("Perlu dicek", len(warns))

    # --- Tab: Unduh | Ringkasan | Peringatan ---
    tab_unduh, tab_ringkasan, tab_warning = st.tabs(
        ["⬇️ Unduh", "📊 Ringkasan section", "⚠️ Perlu dicek"]
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
                file_name=lamp_name,
                mime=mime_docx,
                type="primary",
                use_container_width=True,
            )
            st.caption("Tabel skenario hasil uji fungsional (laporan ASPI).")
        with dl2:
            st.download_button(
                "⬇️ UAT Result (.docx)",
                data=st.session_state.result_docx_bytes,
                file_name=res_name,
                mime=mime_docx,
                type="secondary",
                use_container_width=True,
            )
            st.caption("Dokumen naratif per-case + Daftar Isi.")

    with tab_ringkasan:
        if st.session_state.result_stats is not None:
            used_sections = PRODUCT_PROFILES[used_product]["sections"]
            rows = []
            for section_name, total_baris in used_sections:
                terisi = stats.get(section_name, 0)
                status = "✅ Ditampilkan" if terisi > 0 else "— Tidak dites"
                rows.append(
                    {
                        "Layanan API": section_name,
                        "Baris terisi": terisi,
                        "Status": status,
                    }
                )
            st.dataframe(rows, use_container_width=True, hide_index=True)
        else:
            st.caption("Tidak ada ringkasan.")

    with tab_warning:
        warns = st.session_state.get("result_warnings") or []
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

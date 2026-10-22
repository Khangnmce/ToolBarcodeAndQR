# =============================================================================
# streamlit_app.py - Web App Tao QR Code & Barcode hang loat tu Excel/CSV
# Framework : Streamlit
# Deploy    : Streamlit Cloud (ket noi GitHub repo)
# =============================================================================

import io
import re
import tempfile
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed

import barcode as python_barcode
import pandas as pd
import qrcode
import streamlit as st
from barcode.writer import ImageWriter
from PIL import Image

# ---------------------------------------------------------------------------
# CAU HINH TRANG
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="QR & Barcode Generator",
    page_icon="assets/favicon.png" if __import__("os").path.exists("assets/favicon.png") else ":bar_chart:",
    layout="centered",
    initial_sidebar_state="collapsed",
)


# ---------------------------------------------------------------------------
# CSS TUNG CHINH - DARK THEME HIEN DAI
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    /* Nen trang toi */
    .stApp { background-color: #0f1117; }

    /* Tieu de chinh */
    h1 {
        background: linear-gradient(90deg, #4fc3f7, #1565c0);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.2rem !important;
        font-weight: 800 !important;
        margin-bottom: 0 !important;
    }

    /* Card bao bo cai dat */
    .settings-card {
        background: #1a1d27;
        border: 1px solid #2d3148;
        border-radius: 14px;
        padding: 1.6rem 1.8rem;
        margin-bottom: 1.2rem;
    }

    /* Nhan section */
    .section-label {
        color: #8899aa;
        font-size: 0.78rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 0.4rem;
    }

    /* Nut chinh */
    .stButton > button {
        background: linear-gradient(135deg, #1565c0, #1976d2) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
        padding: 0.65rem 1.4rem !important;
        transition: opacity 0.2s !important;
    }
    .stButton > button:hover { opacity: 0.85 !important; }

    /* Progress bar */
    .stProgress > div > div { background-color: #1976d2 !important; }

    /* Alert success */
    .success-box {
        background: #1b3a2d;
        border-left: 4px solid #43a047;
        border-radius: 8px;
        padding: 0.9rem 1.1rem;
        color: #a5d6a7;
        font-size: 0.95rem;
    }

    /* Alert error */
    .error-box {
        background: #3a1b1b;
        border-left: 4px solid #e53935;
        border-radius: 8px;
        padding: 0.9rem 1.1rem;
        color: #ef9a9a;
        font-size: 0.95rem;
    }

    /* Divider */
    hr { border-color: #2d3148 !important; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ===========================================================================
# HAM LOI: auto_detect_code_column
# Cham diem tung cot de tim cot co kha nang la chua ma code nhat.
# ===========================================================================
def auto_detect_code_column(df: pd.DataFrame) -> str:
    """
    Phan tich DataFrame (mau) va tra ve ten cot co kha nang cao nhat
    la chua ma code dua tren 4 tieu chi thong ke.
    """
    best_col   = None
    best_score = -1.0

    for col in df.columns:
        series = df[col].dropna().astype(str).str.strip()
        series = series[series != ""]
        if len(series) < 2:
            continue

        score = 0.0

        # Tieu chi 1: Alphanumeric thuan (khong dau, khong khoang trang ben trong)
        alphanum_ratio = series.str.match(r"^[A-Za-z0-9\-\_\.]+$").mean()
        score += alphanum_ratio * 40

        # Tieu chi 2: Do dong deu ve do dai
        lengths = series.str.len()
        std_len = lengths.std()
        if std_len <= 2:
            score += 20
        elif std_len <= 5:
            score += 12
        elif std_len <= 10:
            score += 5

        # Tieu chi 3: Do dai trung binh hop ly (4-40 ky tu)
        avg_len = lengths.mean()
        if 4 <= avg_len <= 40:
            score += 20
        elif 2 <= avg_len < 4:
            score += 8

        # Tieu chi 4: Ty le gia tri duy nhat cao
        unique_ratio = series.nunique() / len(series)
        score += unique_ratio * 20

        if score > best_score:
            best_score = score
            best_col   = col

    return best_col


# ===========================================================================
# HAM LOI: render_code_to_bytes
# Sinh anh QR hoac Barcode vao RAM (khong ghi o cung).
# ===========================================================================
def render_code_to_bytes(code: str, mode: str, img_format: str) -> tuple:
    """Tra ve (filename, raw_bytes) cua anh QR hoac Barcode."""
    ext        = img_format.upper()
    pil_format = "JPEG" if ext == "JPG" else "PNG"
    buffer     = io.BytesIO()

    if mode == "QR Code":
        qr = qrcode.QRCode(
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(code)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        img.save(buffer, format=pil_format, optimize=False)
    else:
        writer       = ImageWriter()
        code128_cls  = python_barcode.get_barcode_class("code128")
        code128_obj  = code128_cls(str(code), writer=writer)
        code128_obj.write(buffer, options={"write_text": True, "format": pil_format})

    return f"{code}.{ext.lower()}", buffer.getvalue()


# ===========================================================================
# HAM LOI: build_zip_in_memory
# Doc file, loc ma, sinh anh da luong, dong goi ZIP trong RAM.
# Tra ve bytes cua file ZIP.
# ===========================================================================
def build_zip_in_memory(
    file_bytes: bytes,
    file_ext: str,
    col_name: str,
    mode: str,
    img_format: str,
    progress_callback,          # callback(current, total)
) -> bytes:
    """
    Toan bo qua trinh xu ly chay trong RAM.

    Args:
        file_bytes        : Noi dung file Excel/CSV upload len.
        file_ext          : Duoi file ('xlsx', 'csv', ...).
        col_name          : Ten cot chua ma.
        mode              : Loai ma.
        img_format        : Dinh dang anh.
        progress_callback : Ham cap nhat tien trinh.

    Returns:
        Noi dung file ZIP dang bytes.
    """
    # -- Doc file tu bytes --------------------------------------------------
    buf = io.BytesIO(file_bytes)
    if file_ext in ("xls", "xlsx", "xlsm"):
        df = pd.read_excel(buf, usecols=[col_name])
    else:
        df = pd.read_csv(buf, usecols=[col_name])

    # -- Loc danh sach ma duy nhat ------------------------------------------
    codes = (
        df[col_name]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    )

    if not codes:
        raise ValueError("Danh sach ma rong sau khi loc. Kiem tra lai file va ten cot.")

    total      = len(codes)
    completed  = 0
    compression = (
        zipfile.ZIP_STORED if img_format.upper() == "PNG" else zipfile.ZIP_DEFLATED
    )

    # -- Sinh anh da luong va ghi thang vao ZIP trong RAM -------------------
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, mode="w", compression=compression) as zf:
        with ThreadPoolExecutor(max_workers=16) as executor:
            future_map = {
                executor.submit(render_code_to_bytes, code, mode, img_format): code
                for code in codes
            }
            for future in as_completed(future_map):
                filename, raw_bytes = future.result()
                zf.writestr(filename, raw_bytes)
                completed += 1
                progress_callback(completed, total)

    return zip_buffer.getvalue()


# ===========================================================================
# GIAO DIEN CHINH STREAMLIT
# ===========================================================================

# -- Tieu de & mo ta --------------------------------------------------------
st.markdown("<h1>QR & Barcode Bulk Generator</h1>", unsafe_allow_html=True)
st.markdown(
    "<p style='color:#6b7280; margin-top:0.2rem; margin-bottom:1.5rem;'>"
    "Tao ma hang loat tu file Excel/CSV va tai ve file ZIP</p>",
    unsafe_allow_html=True,
)
st.divider()

# -- Khu vuc upload file ----------------------------------------------------
st.markdown("<div class='section-label'>Buoc 1 — Chon file du lieu</div>", unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    label="Tai len file Excel hoac CSV",
    type=["xlsx", "xls", "xlsm", "csv"],
    label_visibility="collapsed",
)

# -- Khu vuc cai dat --------------------------------------------------------
st.markdown("<div class='section-label' style='margin-top:1.2rem;'>Buoc 2 — Cau hinh</div>", unsafe_allow_html=True)

col_left, col_right = st.columns([1, 1], gap="medium")

with col_left:
    col_name = st.text_input(
        label="Ten cot chua ma",
        value="MaKH",
        help="Ten chinh xac cua cot trong file Excel/CSV chua gia tri ma can tao.",
    )
    # Nut tu dong phat hien cot
    if st.button("Tu dong phat hien cot", use_container_width=True):
        if uploaded_file is None:
            st.warning("Vui long upload file truoc.")
        else:
            try:
                raw = uploaded_file.read()
                uploaded_file.seek(0)   # Reset stream sau khi doc
                ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
                buf = io.BytesIO(raw)
                if ext in ("xls", "xlsx", "xlsm"):
                    sample_df = pd.read_excel(buf, nrows=300)
                else:
                    sample_df = pd.read_csv(buf, nrows=300)

                detected = auto_detect_code_column(sample_df)
                if detected:
                    st.success(f"Phat hien cot: **{detected}**")
                    st.info("Copy ten cot tren vao o 'Ten cot chua ma' phia tren.")
                else:
                    st.warning("Khong phat hien duoc cot phu hop. Vui long nhap thu cong.")
            except Exception as e:
                st.error(f"Loi khi doc file: {e}")

with col_right:
    code_type = st.selectbox(
        label="Loai ma",
        options=["QR Code", "Barcode (Code 128)"],
        index=0,
    )
    img_format = st.selectbox(
        label="Dinh dang anh",
        options=["PNG", "JPG"],
        index=0,
        help="PNG: chat luong tot hon. JPG: nhe hon.",
    )

st.divider()

# -- Nut bat dau ------------------------------------------------------------
st.markdown("<div class='section-label'>Buoc 3 — Tao ma</div>", unsafe_allow_html=True)

if st.button("BAT DAU TAO MA & TAI ZIP", use_container_width=True, type="primary"):
    # Kiem tra dau vao
    if uploaded_file is None:
        st.error("Vui long upload file Excel/CSV truoc.")
    elif not col_name.strip():
        st.error("Vui long nhap ten cot chua ma.")
    else:
        file_ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
        file_bytes = uploaded_file.read()

        # Kiem tra ten cot ton tai truoc khi chay toan bo
        try:
            buf = io.BytesIO(file_bytes)
            if file_ext in ("xls", "xlsx", "xlsm"):
                header_df = pd.read_excel(buf, nrows=0)
            else:
                header_df = pd.read_csv(buf, nrows=0)

            if col_name not in header_df.columns:
                st.error(
                    f"Khong tim thay cot **'{col_name}'** trong file.  \n"
                    f"Cac cot hien co: `{list(header_df.columns)}`"
                )
                st.stop()
        except Exception as e:
            st.error(f"Khong the doc file: {e}")
            st.stop()

        # -- UI tien trinh -------------------------------------------------
        status_text = st.empty()
        progress_bar = st.progress(0)

        def update_progress(current: int, total: int):
            """Cap nhat thanh tien trinh tren Streamlit."""
            pct = int(current / total * 100)
            progress_bar.progress(pct)
            status_text.markdown(
                f"<p style='color:#4fc3f7; font-size:0.9rem;'>"
                f"Dang xu ly {current}/{total} anh... ({pct}%)</p>",
                unsafe_allow_html=True,
            )

        # -- Chay xu ly ----------------------------------------------------
        try:
            status_text.markdown(
                "<p style='color:#8899aa; font-size:0.9rem;'>Dang khoi dong...</p>",
                unsafe_allow_html=True,
            )
            zip_bytes = build_zip_in_memory(
                file_bytes=file_bytes,
                file_ext=file_ext,
                col_name=col_name,
                mode=code_type,
                img_format=img_format,
                progress_callback=update_progress,
            )
            progress_bar.progress(100)
            status_text.markdown(
                "<p style='color:#66bb6a; font-weight:600;'>Hoan thanh!</p>",
                unsafe_allow_html=True,
            )

            # -- Nut download ZIP ------------------------------------------
            output_name = f"output_{code_type.lower().replace(' ', '_').replace('(', '').replace(')', '')}.zip"
            st.download_button(
                label=f"TAI XUONG FILE ZIP ({len(zip_bytes) / 1024:.0f} KB)",
                data=zip_bytes,
                file_name=output_name,
                mime="application/zip",
                use_container_width=True,
            )

        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Loi khong xac dinh: {type(e).__name__}: {e}")

# -- Footer -----------------------------------------------------------------
st.divider()
st.markdown(
    "<p style='text-align:center; color:#374151; font-size:0.8rem;'>"
    "QR & Barcode Bulk Generator &nbsp;|&nbsp; Powered by Streamlit</p>",
    unsafe_allow_html=True,
)

# =============================================================================
# streamlit_app.py - Web App Tao QR Code & Barcode hang loat tu Excel/CSV
# Giao dien: Toi gian trang - den (Monochrome)
# Font chu: Be Vietnam Pro (Google Fonts)
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
# CAU HINH TRANG STREAMLIT
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Tao Ma QR & Barcode Hang Loat",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;0,900;1,400;1,600&display=swap');

    /* Font Be Vietnam Pro toan bo ung dung */
    html, body, [class*="css"], .stApp, h1, h2, h3, h4, p, span, div, button, input, select, label, .stMarkdown {
        font-family: 'Be Vietnam Pro', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif !important;
    }

    /* Nen ung dung: Den sau toi gian */
    .stApp {
        background-color: #09090b !important;
        color: #f4f4f5 !important;
    }

    /* An thanh header mac dinh */
    header[data-testid="stHeader"] {
        background-color: transparent !important;
    }

    /* Tieu de tung buoc */
    .section-title {
        color: #ffffff;
        font-size: 0.88rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        margin-bottom: 0.75rem;
    }

    /* Vung tai tep: Khong vien kep, khong chong cheo */
    [data-testid="stFileUploader"] {
        background: transparent !important;
        border: none !important;
        padding: 0 !important;
    }

    [data-testid="stFileUploadDropzone"],
    [data-testid="stFileUploader"] section {
        background-color: #121214 !important;
        border: 1px dashed #3f3f46 !important;
        border-radius: 8px !important;
        padding: 1.4rem 1.2rem !important;
        transition: border-color 0.2s ease, background-color 0.2s ease;
    }

    [data-testid="stFileUploadDropzone"]:hover,
    [data-testid="stFileUploader"] section:hover {
        border-color: #ffffff !important;
        background-color: #161619 !important;
    }

    [data-testid="stFileUploaderDropzoneInstructions"] {
        display: flex !important;
        flex-direction: row !important;
        align-items: center !important;
        justify-content: space-between !important;
        gap: 1.2rem !important;
        flex-wrap: wrap !important;
    }

    /* Nut Browse files ben trong Dropzone */
    [data-testid="stFileUploadDropzone"] button,
    [data-testid="stFileUploader"] section button {
        background-color: #ffffff !important;
        color: #000000 !important;
        border: 1px solid #ffffff !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        font-size: 0.85rem !important;
        padding: 0.45rem 1.1rem !important;
        cursor: pointer !important;
        transition: all 0.2s ease !important;
    }

    [data-testid="stFileUploadDropzone"] button:hover,
    [data-testid="stFileUploader"] section button:hover {
        background-color: #e4e4e7 !important;
        color: #000000 !important;
        border-color: #e4e4e7 !important;
    }

    /* The hien thi tep da tai len */
    [data-testid="stFileUploaderFile"] {
        background-color: #121214 !important;
        border: 1px solid #27272a !important;
        border-radius: 6px !important;
        padding: 0.6rem 0.8rem !important;
        margin-top: 0.75rem !important;
        color: #f4f4f5 !important;
    }

    /* O nhap lieu text & selectbox */
    .stTextInput input {
        background-color: #121214 !important;
        color: #ffffff !important;
        border: 1px solid #3f3f46 !important;
        border-radius: 6px !important;
        font-weight: 500 !important;
    }
    .stTextInput input:focus {
        border-color: #ffffff !important;
        box-shadow: 0 0 0 1px #ffffff !important;
    }

    [data-baseweb="select"] > div {
        background-color: #121214 !important;
        border: 1px solid #3f3f46 !important;
        border-radius: 6px !important;
        color: #ffffff !important;
    }

    /* Nut bam chinh: Trang tuyet doi chu den */
    .stButton > button {
        background-color: #ffffff !important;
        color: #000000 !important;
        border: 1px solid #ffffff !important;
        border-radius: 6px !important;
        font-weight: 700 !important;
        font-size: 0.92rem !important;
        letter-spacing: 0.02em !important;
        padding: 0.6rem 1.2rem !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button:hover {
        background-color: #e4e4e7 !important;
        color: #000000 !important;
        border-color: #ffffff !important;
        box-shadow: 0 0 12px rgba(255, 255, 255, 0.15) !important;
    }

    /* Nut phu (Tu dong phat hien cot) */
    .secondary-btn button {
        background-color: #121214 !important;
        color: #f4f4f5 !important;
        border: 1px solid #3f3f46 !important;
        font-weight: 500 !important;
        font-size: 0.88rem !important;
        border-radius: 6px !important;
    }
    .secondary-btn button:hover {
        background-color: #1f1f23 !important;
        border-color: #ffffff !important;
        color: #ffffff !important;
    }

    /* Nut tai xuong tap tin ZIP */
    .stDownloadButton > button {
        background-color: #ffffff !important;
        color: #000000 !important;
        border: 1px solid #ffffff !important;
        border-radius: 6px !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
        letter-spacing: 0.02em !important;
        padding: 0.7rem 1.4rem !important;
        box-shadow: 0 0 14px rgba(255, 255, 255, 0.18) !important;
    }
    .stDownloadButton > button:hover {
        background-color: #e4e4e7 !important;
        color: #000000 !important;
        border-color: #ffffff !important;
    }

    /* Thanh tien trinh */
    .stProgress > div > div > div > div {
        background-color: #ffffff !important;
    }
    .stProgress > div > div {
        background-color: #27272a !important;
        border-radius: 4px !important;
        height: 8px !important;
    }

    /* Thong bao trang den */
    .mono-notice {
        background-color: #121214;
        border-left: 3px solid #ffffff;
        border-top: 1px solid #27272a;
        border-right: 1px solid #27272a;
        border-bottom: 1px solid #27272a;
        padding: 0.75rem 1rem;
        border-radius: 6px;
        color: #f4f4f5;
        font-size: 0.88rem;
        margin: 0.75rem 0;
    }

    .mono-alert {
        background-color: #18181b;
        border-left: 3px solid #a1a1aa;
        border-top: 1px solid #27272a;
        border-right: 1px solid #27272a;
        border-bottom: 1px solid #27272a;
        padding: 0.75rem 1rem;
        border-radius: 6px;
        color: #f4f4f5;
        font-size: 0.88rem;
        margin: 0.75rem 0;
    }

    /* Duong phan cach */
    hr {
        border-color: #27272a !important;
        margin: 1.4rem 0 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ===========================================================================
# HAM HO TRO: auto_detect_code_column
# Cham diem tung cot de tim cot co kha nang la chua ma code nhat
# ===========================================================================
def auto_detect_code_column(df: pd.DataFrame) -> str:
    """
    Phan tich DataFrame va tra ve ten cot co kha nang cao nhat
    la chua ma code dua tren 4 tieu chi danh gia.
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
# HAM XU LY: render_code_to_bytes
# Sinh anh QR hoac Barcode truc tiep trong RAM
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
# HAM XU LY: build_zip_in_memory
# Doc du lieu, loc ma, sinh anh song song va dong goi ZIP trong RAM
# ===========================================================================
def build_zip_in_memory(
    file_bytes: bytes,
    file_ext: str,
    col_name: str,
    mode: str,
    img_format: str,
    progress_callback,
) -> bytes:
    buf = io.BytesIO(file_bytes)
    if file_ext in ("xls", "xlsx", "xlsm"):
        df = pd.read_excel(buf, usecols=[col_name])
    else:
        df = pd.read_csv(buf, usecols=[col_name])

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
        raise ValueError("Danh sach ma rong sau khi loc. Vui long kiem tra lai tep va cot da chon.")

    total       = len(codes)
    completed   = 0
    compression = (
        zipfile.ZIP_STORED if img_format.upper() == "PNG" else zipfile.ZIP_DEFLATED
    )

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
# GIAO DIEN CHINH (UI)
# ===========================================================================

# -- TIEU DE UNG DUNG ------------------------------------------------------
st.markdown(
    """
    <div style="margin-top: 0.2rem; margin-bottom: 1.8rem;">
        <h1 style="font-size: 1.65rem; font-weight: 700; color: #ffffff; margin: 0 0 0.35rem 0; letter-spacing: -0.01em;">
            Tạo Mã QR & Barcode Hàng Loạt
        </h1>
        <p style="color: #a1a1aa; font-size: 0.88rem; margin: 0; line-height: 1.5;">
            Đọc dữ liệu từ tệp Excel hoặc CSV, tạo mã tự động và tải về tệp nén ZIP
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# -- BUOC 1: TAI TEP DU LIEU ------------------------------------------------
st.markdown("<div class='section-title'>1. Tải lên tệp dữ liệu</div>", unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    label="Tải lên tệp Excel hoặc CSV",
    type=["xlsx", "xls", "xlsm", "csv"],
    label_visibility="collapsed",
)

st.markdown("<hr/>", unsafe_allow_html=True)

# -- BUOC 2: CAU HINH THONG SO MA -------------------------------------------
st.markdown("<div class='section-title'>2. Cấu hình thông số và cột mã</div>", unsafe_allow_html=True)

col_left, col_right = st.columns([1, 1], gap="large")

with col_left:
    col_name = st.text_input(
        label="Tên cột chứa mã trong bảng tính",
        value="MaKH",
        help="Nhập chính xác tên cột chứa dữ liệu mã code cần tạo.",
    )

    # Nut tu dong phat hien cot
    st.markdown("<div class='secondary-btn'>", unsafe_allow_html=True)
    if st.button("Tự động phát hiện cột mã", use_container_width=True):
        if uploaded_file is None:
            st.markdown(
                "<div class='mono-alert'>Vui lòng tải lên tệp dữ liệu trước khi phát hiện cột.</div>",
                unsafe_allow_html=True,
            )
        else:
            try:
                raw = uploaded_file.read()
                uploaded_file.seek(0)
                ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
                buf = io.BytesIO(raw)
                if ext in ("xls", "xlsx", "xlsm"):
                    sample_df = pd.read_excel(buf, nrows=300)
                else:
                    sample_df = pd.read_csv(buf, nrows=300)

                detected = auto_detect_code_column(sample_df)
                if detected:
                    st.markdown(
                        f"<div class='mono-notice'>Đã phát hiện cột phù hợp: <strong>{detected}</strong>.<br/>"
                        f"Vui lòng điền tên trên vào ô Tên cột nếu chưa trùng khớp.</div>",
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        "<div class='mono-alert'>Không thể tự động nhận diện cột phù hợp. Vui lòng nhập thủ công.</div>",
                        unsafe_allow_html=True,
                    )
            except Exception as e:
                st.markdown(
                    f"<div class='mono-alert'>Lỗi khi đọc tệp dữ liệu: {e}</div>",
                    unsafe_allow_html=True,
                )
    st.markdown("</div>", unsafe_allow_html=True)

with col_right:
    code_type = st.selectbox(
        label="Loại mã cần tạo",
        options=["QR Code", "Barcode (Code 128)"],
        index=0,
    )
    img_format = st.selectbox(
        label="Định dạng hình ảnh xuất",
        options=["PNG", "JPG"],
        index=0,
        help="PNG cho chất lượng sắc nét cao. JPG dung lượng nhẹ hơn.",
    )

st.markdown("<hr/>", unsafe_allow_html=True)

# -- BUOC 3: KHOI TAO MA & TAI TEP NEN --------------------------------------
st.markdown("<div class='section-title'>3. Xử lý và xuất tệp nén</div>", unsafe_allow_html=True)

if st.button("BẮT ĐẦU TẠO MÃ VÀ XUẤT TẬP TIN NÉN", use_container_width=True, type="primary"):
    # Kiem tra dau vao
    if uploaded_file is None:
        st.markdown(
            "<div class='mono-alert'>Yêu cầu tải lên tệp Excel hoặc CSV trước khi tiến hành.</div>",
            unsafe_allow_html=True,
        )
    elif not col_name.strip():
        st.markdown(
            "<div class='mono-alert'>Vui lòng nhập tên cột chứa mã.</div>",
            unsafe_allow_html=True,
        )
    else:
        file_ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
        file_bytes = uploaded_file.read()

        # Kiem tra tieu de cot
        try:
            buf = io.BytesIO(file_bytes)
            if file_ext in ("xls", "xlsx", "xlsm"):
                header_df = pd.read_excel(buf, nrows=0)
            else:
                header_df = pd.read_csv(buf, nrows=0)

            if col_name not in header_df.columns:
                cols_list = ", ".join([f"'{c}'" for c in header_df.columns])
                st.markdown(
                    f"<div class='mono-alert'>"
                    f"Cột '{col_name}' không tồn tại trong tệp.<br/>"
                    f"Danh sách các cột hiện có: {cols_list}."
                    f"</div>",
                    unsafe_allow_html=True,
                )
                st.stop()
        except Exception as e:
            st.markdown(
                f"<div class='mono-alert'>Không thể phân tích dữ liệu tệp: {e}</div>",
                unsafe_allow_html=True,
            )
            st.stop()

        # Tien trinh thuc hien
        status_text = st.empty()
        progress_bar = st.progress(0)

        def update_progress(current: int, total: int):
            pct = int(current / total * 100)
            progress_bar.progress(pct)
            status_text.markdown(
                f"<div style='font-size:0.88rem; color:#d4d4d8; font-weight:500; margin-bottom:0.4rem;'>"
                f"Tiến độ xử lý: {current} trên {total} hình ảnh ({pct}%)"
                f"</div>",
                unsafe_allow_html=True,
            )

        try:
            status_text.markdown(
                "<div style='font-size:0.88rem; color:#a1a1aa; margin-bottom:0.4rem;'>Đang khởi tạo các tiến trình nền...</div>",
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
                "<div class='mono-notice'>Quá trình tạo mã đã hoàn tất thành công.</div>",
                unsafe_allow_html=True,
            )

            # Nut tai tap tin ZIP
            clean_type = "qr" if code_type == "QR Code" else "barcode"
            output_name = f"ma_{clean_type}_{col_name}.zip"
            st.download_button(
                label=f"TẢI XUỐNG TỆP NÉN ZIP ({len(zip_bytes) / 1024:.0f} KB)",
                data=zip_bytes,
                file_name=output_name,
                mime="application/zip",
                use_container_width=True,
            )

        except ValueError as e:
            st.markdown(
                f"<div class='mono-alert'>{str(e)}</div>",
                unsafe_allow_html=True,
            )
        except Exception as e:
            st.markdown(
                f"<div class='mono-alert'>Lỗi trong quá trình xử lý: {type(e).__name__} - {e}</div>",
                unsafe_allow_html=True,
            )


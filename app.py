# =============================================================================
# app.py - Cong cu Tao QR Code & Barcode hang loat tu Excel/CSV
# Tac gia  : Senior Python Developer
# Thu vien : customtkinter, pandas, qrcode, python-barcode, Pillow
# Mo ta    : Doc danh sach ma tu file Excel/CSV, sinh anh QR/Barcode da luong
#             (khong ghi ra o cung), roi dong goi toan bo vao mot file ZIP.
# =============================================================================

import io
import threading
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from tkinter import filedialog, messagebox

import barcode as python_barcode
import customtkinter as ctk
import pandas as pd
import qrcode
from barcode.writer import ImageWriter
from PIL import Image

# ---------------------------------------------------------------------------
# CAU HINH GIAO DIEN TOAN CUC
# ---------------------------------------------------------------------------
ctk.set_appearance_mode("Dark")          # Giao dien toi (Dark Mode)
ctk.set_default_color_theme("blue")      # Chu de mau xanh


# ===========================================================================
# HAM LOI: render_code_to_bytes
# Sinh anh QR Code hoac Barcode Code128 truc tiep vao bo nho (RAM),
# KHONG ghi bat ky file nao ra o cung.
# Tra ve: tuple (ten_file: str, du_lieu_anh: bytes)
# ===========================================================================
def render_code_to_bytes(code: str, mode: str, img_format: str) -> tuple:
    """
    Tao anh ma QR hoac Barcode va tra ve duoi dang bytes trong RAM.

    Args:
        code       : Chuoi du lieu can ma hoa.
        mode       : 'QR Code' hoac 'Barcode (Code 128)'.
        img_format : 'PNG' hoac 'JPG'.

    Returns:
        (filename, raw_bytes) - ten file va noi dung nhi phan cua anh.
    """
    ext = img_format.upper()                        # 'PNG' hoac 'JPG'
    pil_format = "JPEG" if ext == "JPG" else "PNG"  # Pillow dung 'JPEG'

    buffer = io.BytesIO()  # Vung dem trong RAM - khong dungden o cung

    if mode == "QR Code":
        # -- Tao QR Code --------------------------------------------------
        qr = qrcode.QRCode(
            error_correction=qrcode.constants.ERROR_CORRECT_L,  # Toi uu toc do
            box_size=10,
            border=4,
        )
        qr.add_data(code)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        img.save(buffer, format=pil_format, optimize=False)

    else:
        # -- Tao Barcode Code 128 -----------------------------------------
        writer = ImageWriter()
        code128_class = python_barcode.get_barcode_class("code128")
        code128_obj = code128_class(str(code), writer=writer)

        # Ghi barcode vao buffer; write_text=True hien thi so ben duoi vach
        code128_obj.write(
            buffer,
            options={
                "write_text": True,
                "format": pil_format,
            },
        )

    filename = f"{code}.{ext.lower()}"
    return filename, buffer.getvalue()


# ===========================================================================
# HAM HO TRO: auto_detect_code_column
# Phan tich mau du lieu cua tung cot va chon cot co kha nang cao nhat
# la chua ma code (barcode / QR), dua tren cac tieu chi:
#   1. Ty le ky tu alphanumeric thuan tuy (khong dau, khong khoang trang ben trong)
#   2. Do dong deu ve do dai (do lech chuan thap)
#   3. Do dai trung binh hop ly cho ma code (4 - 40 ky tu)
#   4. Ty le gia tri duy nhat cao (ma it trung lap)
# ===========================================================================
def auto_detect_code_column(df: pd.DataFrame) -> str:
    """
    Tu dong phat hien cot chua ma code dua tren dac diem thong ke.

    Args:
        df: DataFrame da doc tu file Excel/CSV (mau toi da 300 dong).

    Returns:
        Ten cot duoc danh gia la chua ma code, hoac None neu khong phat hien.
    """
    import re

    best_col   = None
    best_score = -1.0

    for col in df.columns:
        series = df[col].dropna().astype(str).str.strip()
        series = series[series != ""]  # Bo gia tri rong
        if len(series) < 2:
            continue

        score = 0.0

        # -- Tieu chi 1: Ty le ky tu alphanumeric thuan (khong dau tieng Viet,
        #    khong khoang trang ben trong, chi cho phep A-Z, 0-9, - _ .)
        #    Diem toi da: 40
        alphanum_ratio = series.str.match(r'^[A-Za-z0-9\-\_\.]+$').mean()
        score += alphanum_ratio * 40

        # -- Tieu chi 2: Do dong deu ve do dai (std thap = ma co dinh hoac ngan)
        #    Diem toi da: 20
        lengths = series.str.len()
        std_len = lengths.std()
        if std_len <= 2:
            score += 20
        elif std_len <= 5:
            score += 12
        elif std_len <= 10:
            score += 5

        # -- Tieu chi 3: Do dai trung binh nam trong khoang hop ly
        #    Ma code thong thuong: 4 - 40 ky tu
        #    Diem toi da: 20
        avg_len = lengths.mean()
        if 4 <= avg_len <= 40:
            score += 20
        elif 2 <= avg_len < 4:
            score += 8

        # -- Tieu chi 4: Ty le gia tri duy nhat cao (ma hiem khi bi trung lap)
        #    Diem toi da: 20
        unique_ratio = series.nunique() / len(series)
        score += unique_ratio * 20

        if score > best_score:
            best_score = score
            best_col   = col

    return best_col


# ===========================================================================
# HAM LOI: process_and_zip
# Ham chay trong luong nen (background thread).
# Doc Excel/CSV -> sinh anh da luong -> ghi ZIP tung anh ngay khi hoan tat.
# ===========================================================================
def process_and_zip(
    excel_path,
    col_name,
    mode,
    img_format,
    zip_path,
    on_progress,
    on_done,
):
    """
    Luong xu ly chinh: doc du lieu -> song song hoa render -> ghi ZIP.

    Args:
        excel_path : Duong dan file Excel/CSV dau vao.
        col_name   : Ten cot chua ma trong file.
        mode       : Loai ma ('QR Code' hoac 'Barcode (Code 128)').
        img_format : Dinh dang anh ('PNG' hoac 'JPG').
        zip_path   : Duong dan file ZIP dau ra.
        on_progress: Ham callback cap nhat tien trinh (thread-safe qua after()).
        on_done    : Ham callback thong bao hoan thanh hoac loi.
    """
    try:
        # -- 1. Doc file Excel hoac CSV tuy theo duoi file -----------------
        ext = excel_path.rsplit(".", 1)[-1].lower()
        if ext in ("xls", "xlsx", "xlsm"):
            # usecols: chi doc cot can thiet, bo qua cac cot thua -> nhanh hon
            df = pd.read_excel(excel_path, usecols=[col_name])
        elif ext == "csv":
            # usecols: tuong tu, giam thoi gian parse va RAM voi file lon
            df = pd.read_csv(excel_path, usecols=[col_name])
        else:
            raise ValueError(f"Dinh dang file khong duoc ho tro: .{ext}")

        # -- 2. Lay cot ma; bat KeyError neu ten cot sai -------------------
        if col_name not in df.columns:
            raise KeyError(
                f"Khong tim thay cot '{col_name}' trong file.\n"
                f"Cac cot hien co: {list(df.columns)}"
            )

        # Lay danh sach ma duy nhat, loai bo NaN va chuoi rong/khoang trang
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
            raise ValueError(
                "Danh sach ma rong sau khi loc. Kiem tra lai file va ten cot."
            )

        total = len(codes)

        # -- 3. Chon thuat toan nen ZIP ------------------------------------
        # PNG da nen san -> ZIP_STORED (nhanh hon, khong lang phi CPU).
        # JPG cung da nen nhung ZIP_DEFLATED doi khi nho hon mot chut.
        compression = (
            zipfile.ZIP_STORED
            if img_format.upper() == "PNG"
            else zipfile.ZIP_DEFLATED
        )

        # -- 4. Mo file ZIP va chay ThreadPoolExecutor ---------------------
        completed = 0
        with zipfile.ZipFile(zip_path, mode="w", compression=compression) as zf:
            with ThreadPoolExecutor(max_workers=16) as executor:
                # Gui tat ca tac vu vao pool
                future_map = {
                    executor.submit(render_code_to_bytes, code, mode, img_format): code
                    for code in codes
                }

                # Lang nghe tung future hoan tat; ghi ngay vao ZIP
                for future in as_completed(future_map):
                    filename, raw_bytes = future.result()  # Raise neu loi ben trong
                    zf.writestr(filename, raw_bytes)        # Ghi vao ZIP truc tiep

                    completed += 1
                    on_progress(completed, total)           # Cap nhat UI qua app.after()

        on_done(True, f"Hoan thanh! Da tao {total} anh.\nFile ZIP: {zip_path}")

    except KeyError as e:
        on_done(False, f"Loi ten cot:\n{e}")
    except FileNotFoundError:
        on_done(False, f"Khong tim thay file:\n{excel_path}")
    except Exception as e:
        on_done(False, f"Loi khong xac dinh:\n{type(e).__name__}: {e}")


# ===========================================================================
# LOP GIAO DIEN CHINH: App
# ===========================================================================
class App(ctk.CTk):
    """Cua so ung dung chinh."""

    # -- Hang so giao dien -------------------------------------------------
    WINDOW_TITLE  = "QR & Barcode Generator  v1.0"
    WINDOW_WIDTH  = 680
    WINDOW_HEIGHT = 480
    PAD_X         = 20
    PAD_Y         = 10

    def __init__(self):
        super().__init__()

        # -- Cau hinh cua so -----------------------------------------------
        self.title(self.WINDOW_TITLE)
        self.geometry(f"{self.WINDOW_WIDTH}x{self.WINDOW_HEIGHT}")
        self.resizable(False, False)

        # -- Bien trang thai -----------------------------------------------
        self._excel_path = ""     # Duong dan file Excel/CSV da chon
        self._is_running = False  # Co ngan chay dong thoi nhieu tien trinh

        # -- Xay dung giao dien --------------------------------------------
        self._build_ui()

    # -----------------------------------------------------------------------
    # XAY DUNG GIAO DIEN
    # -----------------------------------------------------------------------
    def _build_ui(self):
        """Khoi tao va bo cuc toan bo widget."""

        # -- Tieu de ung dung ----------------------------------------------
        lbl_title = ctk.CTkLabel(
            self,
            text="QR & Barcode Bulk Generator",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        lbl_title.pack(pady=(20, 5))

        lbl_sub = ctk.CTkLabel(
            self,
            text="Tao ma hang loat tu Excel/CSV va xuat ra file ZIP",
            font=ctk.CTkFont(size=12),
            text_color="gray",
        )
        lbl_sub.pack(pady=(0, 15))

        # -- Frame chua tat ca input ---------------------------------------
        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.pack(padx=self.PAD_X, fill="x")

        # -- Hang 1: Chon file Excel/CSV -----------------------------------
        row1 = ctk.CTkFrame(frame, fg_color="transparent")
        row1.pack(padx=self.PAD_X, pady=(self.PAD_Y + 5, self.PAD_Y), fill="x")

        btn_file = ctk.CTkButton(
            row1,
            text="Chon file Excel / CSV",
            width=200,
            command=self._browse_file,
        )
        btn_file.pack(side="left")

        self.lbl_file_path = ctk.CTkLabel(
            row1,
            text="Chua chon file...",
            text_color="gray",
            anchor="w",
            wraplength=380,
        )
        self.lbl_file_path.pack(side="left", padx=(12, 0), fill="x", expand=True)

        # -- Hang 2: Ten cot chua ma ---------------------------------------
        row2 = ctk.CTkFrame(frame, fg_color="transparent")
        row2.pack(padx=self.PAD_X, pady=self.PAD_Y, fill="x")

        lbl_col = ctk.CTkLabel(row2, text="Ten cot chua ma:", width=160, anchor="w")
        lbl_col.pack(side="left")

        self.entry_col = ctk.CTkEntry(
            row2,
            placeholder_text="VD: MaKH",
            width=220,
        )
        self.entry_col.insert(0, "MaKH")   # Gia tri mac dinh
        self.entry_col.pack(side="left", padx=(8, 0))

        # Nut tu dong phat hien cot ma code
        self.btn_auto = ctk.CTkButton(
            row2,
            text="Tu dong",
            width=90,
            fg_color="#2d6a4f",
            hover_color="#1b4332",
            command=self._auto_detect_column,
        )
        self.btn_auto.pack(side="left", padx=(10, 0))

        # -- Hang 3: Loai ma & Dinh dang anh -------------------------------
        row3 = ctk.CTkFrame(frame, fg_color="transparent")
        row3.pack(padx=self.PAD_X, pady=self.PAD_Y, fill="x")

        lbl_type = ctk.CTkLabel(row3, text="Loai ma:", width=100, anchor="w")
        lbl_type.pack(side="left")

        self.combo_type = ctk.CTkComboBox(
            row3,
            values=["QR Code", "Barcode (Code 128)"],
            width=200,
            state="readonly",
        )
        self.combo_type.set("QR Code")
        self.combo_type.pack(side="left", padx=(8, 24))

        lbl_fmt = ctk.CTkLabel(row3, text="Dinh dang:", width=90, anchor="w")
        lbl_fmt.pack(side="left")

        self.combo_format = ctk.CTkComboBox(
            row3,
            values=["PNG", "JPG"],
            width=100,
            state="readonly",
        )
        self.combo_format.set("PNG")
        self.combo_format.pack(side="left", padx=(8, 0))

        # -- Separator -----------------------------------------------------
        sep = ctk.CTkFrame(frame, height=1, fg_color="#3a3a3a")
        sep.pack(padx=self.PAD_X, pady=(self.PAD_Y, 0), fill="x")

        # -- Hang 4: Nut BAT DAU -------------------------------------------
        row4 = ctk.CTkFrame(frame, fg_color="transparent")
        row4.pack(padx=self.PAD_X, pady=(self.PAD_Y + 2, self.PAD_Y + 5), fill="x")

        self.btn_start = ctk.CTkButton(
            row4,
            text="BAT DAU TAO MA & XUAT ZIP",
            height=42,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#1f6aa5",
            hover_color="#144d78",
            command=self._on_start,
        )
        self.btn_start.pack(fill="x")

        # -- Khu vuc tien trinh --------------------------------------------
        prog_frame = ctk.CTkFrame(self, fg_color="transparent")
        prog_frame.pack(padx=self.PAD_X, pady=(14, 0), fill="x")

        self.progressbar = ctk.CTkProgressBar(prog_frame, height=14, corner_radius=7)
        self.progressbar.set(0)
        self.progressbar.pack(fill="x")

        # -- Nhan trang thai -----------------------------------------------
        self.lbl_status = ctk.CTkLabel(
            self,
            text="Dang cho...",
            font=ctk.CTkFont(size=12),
            text_color="gray",
        )
        self.lbl_status.pack(pady=(8, 0))

    # -----------------------------------------------------------------------
    # XU LY SU KIEN: Chon file
    # -----------------------------------------------------------------------
    def _browse_file(self):
        """Mo hop thoai chon file Excel hoac CSV."""
        path = filedialog.askopenfilename(
            title="Chon file du lieu",
            filetypes=[
                ("Excel & CSV files", "*.xlsx *.xls *.xlsm *.csv"),
                ("Excel files", "*.xlsx *.xls *.xlsm"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self._excel_path = path
            # Hien thi ten file ngan gon (chi phan ten + duoi)
            short = path.replace("\\", "/").split("/")[-1]
            self.lbl_file_path.configure(text=short, text_color="#4fc3f7")

    # -----------------------------------------------------------------------
    # XU LY SU KIEN: Tu dong phat hien cot
    # -----------------------------------------------------------------------
    def _auto_detect_column(self):
        """
        Doc mau file da chon (toi da 300 dong), goi auto_detect_code_column()
        de tim cot co kha nang cao nhat la chua ma code, sau do dien vao entry.
        """
        if not self._excel_path:
            messagebox.showwarning(
                "Chua chon file",
                "Vui long chon file Excel hoac CSV truoc khi tu dong phat hien cot.",
            )
            return

        try:
            ext = self._excel_path.rsplit(".", 1)[-1].lower()
            # Chi doc mau 300 dong dau de phan tich nhanh, khong can ca file
            if ext in ("xls", "xlsx", "xlsm"):
                sample_df = pd.read_excel(self._excel_path, nrows=300)
            elif ext == "csv":
                sample_df = pd.read_csv(self._excel_path, nrows=300)
            else:
                messagebox.showerror("Loi", f"Dinh dang khong ho tro: .{ext}")
                return

            detected = auto_detect_code_column(sample_df)

            if detected is None:
                messagebox.showwarning(
                    "Khong phat hien",
                    "Khong tim thay cot phu hop voi dang ma code.\n"
                    "Vui long nhap ten cot thu cong.",
                )
                return

            # Dien ten cot phat hien duoc vao o nhap
            self.entry_col.delete(0, "end")
            self.entry_col.insert(0, detected)
            self.lbl_status.configure(
                text=f"Tu dong phat hien cot: '{detected}'",
                text_color="#4fc3f7",
            )

        except Exception as e:
            messagebox.showerror("Loi", f"Khong the doc file:\n{e}")

    # -----------------------------------------------------------------------
    # XU LY SU KIEN: Bat dau
    # -----------------------------------------------------------------------
    def _on_start(self):
        """Kiem tra dau vao, mo hop thoai luu ZIP, roi khoi chay luong nen."""

        # -- Kiem tra nguoi dung da chon file chua -------------------------
        if not self._excel_path:
            messagebox.showwarning(
                "Thieu file dau vao",
                "Vui long chon file Excel hoac CSV truoc khi bat dau.",
            )
            return

        # -- Lay ten cot; kiem tra khong rong ------------------------------
        col_name = self.entry_col.get().strip()
        if not col_name:
            messagebox.showwarning(
                "Thieu ten cot",
                "Vui long nhap ten cot chua ma.",
            )
            return

        # -- Mo hop thoai chon noi luu file ZIP ----------------------------
        zip_path = filedialog.asksaveasfilename(
            title="Luu file ZIP ket qua",
            defaultextension=".zip",
            filetypes=[("ZIP archive", "*.zip")],
            initialfile="output_codes.zip",
        )
        if not zip_path:
            return  # Nguoi dung bam Cancel -> khong lam gi

        # -- Lay cac tham so con lai ---------------------------------------
        mode       = self.combo_type.get()
        img_format = self.combo_format.get()

        # -- Reset giao dien, vo hieu hoa nut Bat dau ----------------------
        self._set_running(True)
        self.progressbar.set(0)
        self.lbl_status.configure(text="Dang khoi dong...", text_color="white")

        # -- Khoi chay luong nen - KHONG lam do UI -------------------------
        worker = threading.Thread(
            target=process_and_zip,
            args=(
                self._excel_path,
                col_name,
                mode,
                img_format,
                zip_path,
                self._on_progress_callback,
                self._on_done_callback,
            ),
            daemon=True,  # Luong tu tat khi cua so dong
        )
        worker.start()

    # -----------------------------------------------------------------------
    # CALLBACK THREAD-SAFE: Cap nhat tien trinh
    # -----------------------------------------------------------------------
    def _on_progress_callback(self, current: int, total: int):
        """
        Duoc goi tu luong nen; dung app.after(0, ...) de day cap nhat
        vao hang doi su kien cua Tkinter - an toan voi da luong.
        """
        self.after(0, self._update_progress_ui, current, total)

    def _update_progress_ui(self, current: int, total: int):
        """Cap nhat ProgressBar va Label trang thai (chay tren main thread)."""
        progress = current / total if total > 0 else 0
        self.progressbar.set(progress)
        self.lbl_status.configure(
            text=f"Dang xu ly {current}/{total} anh... ({progress:.0%})",
            text_color="white",
        )

    # -----------------------------------------------------------------------
    # CALLBACK THREAD-SAFE: Hoan thanh hoac loi
    # -----------------------------------------------------------------------
    def _on_done_callback(self, success: bool, message: str):
        """
        Duoc goi tu luong nen khi qua trinh ket thuc (thanh cong hoac loi).
        Dung app.after(0, ...) de dam bao an toan Tkinter.
        """
        self.after(0, self._finalize_ui, success, message)

    def _finalize_ui(self, success: bool, message: str):
        """Cap nhat UI cuoi cung va hien thi popup thong bao (main thread)."""
        self._set_running(False)

        if success:
            self.progressbar.set(1.0)
            self.lbl_status.configure(
                text="Hoan thanh!",
                text_color="#66bb6a",
            )
            messagebox.showinfo("Thanh cong", message)
        else:
            self.progressbar.set(0)
            self.lbl_status.configure(
                text="Da xay ra loi. Kiem tra thong bao.",
                text_color="#ef5350",
            )
            messagebox.showerror("Loi", message)

    # -----------------------------------------------------------------------
    # TIEN ICH: Bat/Tat trang thai dang chay
    # -----------------------------------------------------------------------
    def _set_running(self, running: bool):
        """Vo hieu hoa / kich hoat lai nut Bat dau khi tien trinh chay."""
        self._is_running = running
        state = "disabled" if running else "normal"
        self.btn_start.configure(state=state)


# ===========================================================================
# DIEM KHOI CHAY CHUONG TRINH
# ===========================================================================
if __name__ == "__main__":
    app = App()
    app.mainloop()

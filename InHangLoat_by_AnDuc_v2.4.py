"""
In Hàng Loạt by An Duc v2.4
Công cụ in hàng loạt file Excel (.xls/.xlsx) và PDF (.pdf) trong 1 thư mục.
"""

import os
import gc
import re
import tempfile
import threading
import queue
import time
import uuid
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import win32com.client
import win32print
import pythoncom
from pypdf import PdfReader, PdfWriter

TEN_PROJECT = "In Hàng Loạt by An Duc v2.4"


# Hằng số hướng giấy của Windows (winspool.h): 1 = dọc (Portrait), 2 = ngang (Landscape).
DMORIENT_PORTRAIT = 1
DMORIENT_LANDSCAPE = 2


def dat_cau_hinh_may_in_mac_dinh(so_ban: int, la_ngang: bool):
    """
    Đặt CẢ số bản in ("Copies") LẪN hướng giấy (dọc/ngang) thẳng vào driver của
    máy in MẶC ĐỊNH trên Windows, trong CÙNG 1 lần mở/đóng driver (thay vì 2 lần
    riêng biệt) — gọn hơn và giảm số lần truy cập driver không cần thiết.

    LÝ DO CẦN HÀM NÀY: các cách in PDF im lặng (verb "print" của Windows) đều
    KHÔNG có tham số riêng để chỉ định số bản in hay hướng giấy — chúng luôn in
    theo đúng cấu hình đang được LƯU SẴN trong driver máy in mặc định (y hệt
    các ô "Copies"/"Orientation" trong hộp thoại in tay). Nếu cấu hình đó đang
    "dính" giá trị khác từ lần in tay trước đó (của người dùng hoặc bất kỳ phần
    mềm nào khác từng in trên máy đó), MỌI lệnh in im lặng sau này sẽ tự động
    bị áp sai theo cấu hình đó — đây chính là nguyên nhân gây in ra nhiều bản
    hơn mong muốn, hoặc in sai hướng giấy khiến nội dung bị co nhỏ lệch góc.

    Excel không gặp vấn đề số bản in vì Excel COM có tham số Copies riêng, tự
    kiểm soát độc lập với driver.

    Trả về True nếu đặt thành công, False nếu thất bại (ví dụ không tìm thấy
    máy in mặc định) — khi thất bại, code gọi hàm này nên vẫn tiếp tục gửi lệnh
    in bình thường (chấp nhận rủi ro số bản/hướng giấy có thể không chính xác)
    thay vì dừng hẳn, vì đây là tính năng "cố gắng tối ưu", không phải điều
    kiện bắt buộc.
    """
    try:
        ten_may_in = win32print.GetDefaultPrinter()
        handle = win32print.OpenPrinter(ten_may_in)
        try:
            thong_tin = win32print.GetPrinter(handle, 2)
            devmode = thong_tin.get("pDevMode")
            if devmode is None:
                return False
            devmode.Copies = so_ban
            devmode.Orientation = DMORIENT_LANDSCAPE if la_ngang else DMORIENT_PORTRAIT
            thong_tin["pDevMode"] = devmode
            win32print.SetPrinter(handle, 2, thong_tin, 0)
            return True
        finally:
            win32print.ClosePrinter(handle)
    except Exception:
        return False





def kiem_tra_pdf_huong_ngang(trang) -> bool:
    """
    Kiểm tra 1 trang PDF (đối tượng page của pypdf) có phải hướng ngang
    (landscape) hay không, dựa vào kích thước MediaBox và góc xoay (/Rotate)
    của trang — vì nếu trang bị xoay 90/270 độ, chiều rộng/cao thực tế hiển
    thị bị hoán đổi so với MediaBox gốc.
    Trả về False (coi như hướng dọc) nếu không xác định được vì bất kỳ lý do gì.
    """
    try:
        goc_xoay = int(trang.get("/Rotate") or 0) % 360
        hop = trang.mediabox
        rong, cao = float(hop.width), float(hop.height)
        if goc_xoay in (90, 270):
            rong, cao = cao, rong
        return rong > cao
    except Exception:
        return False


def _lay_danh_sach_job_id_dang_cho(ten_may_in=None):
    """
    Trả về TẬP HỢP các Job ID đang có trong hàng đợi của máy in (mặc định nếu
    không truyền tên). Trả về None nếu không lấy được (ví dụ không truy cập
    được máy in) — dùng làm cờ báo "không tin cậy được, đừng dựa vào kết quả này".
    """
    try:
        if ten_may_in is None:
            ten_may_in = win32print.GetDefaultPrinter()
        handle = win32print.OpenPrinter(ten_may_in)
        try:
            danh_sach = win32print.EnumJobs(handle, 0, 1000, 1)
            return {job["JobId"] for job in danh_sach}
        finally:
            win32print.ClosePrinter(handle)
    except Exception:
        return None


def doi_lenh_in_duoc_gui_vao_hang_doi(danh_sach_id_truoc_khi_gui, thoi_gian_cho_toi_da: float = 10.0) -> bool:
    """
    Đợi cho đến khi hàng đợi máy in xuất hiện 1 Job ID HOÀN TOÀN MỚI (so với
    tập hợp ID đang có trước khi gửi lệnh), để biết chắc ứng dụng đọc PDF đã
    đọc xong file và nộp lệnh in vào hệ thống — thay vì chỉ đợi 1 khoảng thời
    gian cố định (không đáng tin cậy) hoặc chỉ so sánh SỐ LƯỢNG job đơn thuần.

    QUAN TRỌNG — LÝ DO PHẢI SO SÁNH THEO ID CHỨ KHÔNG PHẢI SỐ LƯỢNG: nếu job
    in của file TRƯỚC vừa in xong và bị xóa khỏi hàng đợi ĐÚNG LÚC job in của
    file HIỆN TẠI vừa được thêm vào, tổng SỐ LƯỢNG job có thể không đổi (1 job
    cũ rời đi, 1 job mới đến cùng lúc) — khiến cách đếm số lượng đơn thuần (bản
    trước) bị bỏ lỡ, tưởng nhầm là chưa có gì thay đổi dù job mới đã thực sự
    được gửi, khiến chương trình phải đợi hết timeout một cách oan uổng. Với
    các file PDF nhỏ (in rất nhanh), tình huống này xảy ra khá thường xuyên.
    So sánh theo ID cụ thể (mỗi job luôn được Windows cấp 1 ID mới, không bao
    giờ trùng lặp/tái sử dụng) tránh được hoàn toàn vấn đề này.

    LƯU Ý VỀ GIỚI HẠN: hàm này chỉ xác nhận được lệnh in đã được HỆ ĐIỀU HÀNH
    chấp nhận vào hàng đợi máy in (uplift đáng kể so với chỉ "đã gửi yêu cầu
    mở app"), nhưng KHÔNG thể xác nhận việc in vật lý đã hoàn tất thật sự (mực
    đã ra giấy) — điều đó nằm ngoài khả năng theo dõi ở tầng phần mềm này. Với
    máy in xử lý cực nhanh, job có thể xuất hiện rồi biến mất hoàn toàn giữa
    2 lần kiểm tra (chu kỳ 0.15 giây) — trường hợp cực hiếm đó vẫn được xử lý
    an toàn nhờ khoảng chờ tối đa (timeout) đóng vai trò như một khoảng chờ cố
    định dự phòng.

    Trả về True nếu phát hiện Job ID mới trong hàng đợi, False nếu hết thời
    gian chờ mà không phát hiện được (không có nghĩa là lệnh in thất bại, chỉ
    là không xác nhận được — chương trình vẫn tiếp tục xử lý bình thường).
    """
    if danh_sach_id_truoc_khi_gui is None:
        # Không lấy được danh sách ngay từ đầu (ví dụ không lấy được máy in mặc
        # định) — dùng khoảng chờ cố định làm phương án dự phòng duy nhất còn lại.
        time.sleep(1.5)
        return False

    thoi_diem_bat_dau = time.time()
    while time.time() - thoi_diem_bat_dau < thoi_gian_cho_toi_da:
        danh_sach_hien_tai = _lay_danh_sach_job_id_dang_cho()
        if danh_sach_hien_tai is None:
            break  # mất khả năng truy vấn giữa chừng, dừng đợi, coi như không xác nhận được
        if danh_sach_hien_tai - danh_sach_id_truoc_khi_gui:
            # Có ít nhất 1 ID mới xuất hiện mà trước đó không có.
            return True
        time.sleep(0.15)
    return False


def parse_danh_sach_so(chuoi: str, ten_de_bao_loi: str = "danh sách") -> list[int]:
    """
    Phân tích chuỗi nhập của người dùng thành danh sách số nguyên, giữ đúng thứ tự
    người dùng nhập, không trùng lặp. Hỗ trợ:
      "1,2"      -> [1, 2]
      "3-4"      -> [3, 4]
      "1-2,4"    -> [1, 2, 4]
      "1,3-4,2"  -> [1, 3, 4, 2]  (giữ đúng thứ tự người dùng nhập)
    Ném ValueError nếu chuỗi không hợp lệ hoặc rỗng.
    """
    chuoi = chuoi.strip()
    if not chuoi:
        raise ValueError(f"{ten_de_bao_loi} đang trống.")

    ket_qua: list[int] = []
    da_them: set[int] = set()

    for phan in chuoi.split(","):
        phan = phan.strip()
        if not phan:
            continue
        if "-" in phan:
            a, b = phan.split("-", 1)
            bat_dau, ket_thuc = int(a.strip()), int(b.strip())
            if bat_dau < 1 or ket_thuc < bat_dau:
                raise ValueError(f"Khoảng '{phan}' không hợp lệ.")
            for so in range(bat_dau, ket_thuc + 1):
                if so not in da_them:
                    ket_qua.append(so)
                    da_them.add(so)
        else:
            so = int(phan)
            if so < 1:
                raise ValueError(f"Số '{phan}' không hợp lệ.")
            if so not in da_them:
                ket_qua.append(so)
                da_them.add(so)

    if not ket_qua:
        raise ValueError(f"{ten_de_bao_loi} đang trống.")
    return ket_qua


def gom_nhom_lien_tuc(danh_sach_so: list[int]) -> list[tuple[int, int]]:
    """
    Gom các số liền kề nhau thành từng khoảng (bắt_đầu, kết_thúc) để giảm số lệnh in
    thực tế phải gửi. VD: [1, 2, 3, 5, 7, 8] -> [(1, 3), (5, 5), (7, 8)]
    """
    if not danh_sach_so:
        return []
    danh_sach_sap_xep = sorted(set(danh_sach_so))
    ket_qua = []
    bat_dau = ket_thuc_truoc = danh_sach_sap_xep[0]
    for so in danh_sach_sap_xep[1:]:
        if so == ket_thuc_truoc + 1:
            ket_thuc_truoc = so
            continue
        ket_qua.append((bat_dau, ket_thuc_truoc))
        bat_dau = ket_thuc_truoc = so
    ket_qua.append((bat_dau, ket_thuc_truoc))
    return ket_qua


def khoa_sap_xep_tu_nhien(chuoi: str):
    """
    Tạo khóa sắp xếp "tự nhiên": so sánh phần số theo giá trị số thực (2 < 10),
    thay vì so sánh chuỗi thuần túy (vốn sẽ xếp "10" trước "2" vì ký tự '1' < '2').
    VD sắp xếp bằng khóa này: "1_a.xlsx", "2_a.xlsx", "10_a.xlsx" (đúng thứ tự số).
    """
    return [int(phan) if phan.isdigit() else phan.lower() for phan in re.split(r"(\d+)", chuoi)]


class InPageApp:
    def __init__(self):
        self.thu_muc_goc = None
        self.danh_sach_file = []
        self.stop_event = threading.Event()
        self.log_queue = queue.Queue()

        # Loại file cần in: "excel" hoặc "pdf". Chỉ chọn 1 trong 2 — khi đã chọn,
        # chỉ file đúng loại đó được quét/in, loại còn lại bị bỏ qua hoàn toàn và
        # màn cài đặt chỉ hiện đúng các tùy chọn liên quan tới loại đã chọn.
        self.loai_file = None

        # Cấu hình chọn sheet cần in (chỉ áp dụng khi loai_file == "excel"):
        #   "vi_tri"    -> in theo MỘT số thứ tự sheet (1 = sheet đầu tiên, 2 = sheet thứ 2, ...)
        #   "danh_sach" -> in theo danh sách/khoảng số thứ tự sheet, vd "1,2" hoặc "3-4" hoặc "1-2,4"
        #   "ten"       -> in theo đúng tên sheet (áp dụng cho mọi file, ví dụ "Sheet2")
        #   "tat_ca"    -> in tất cả các sheet có trong file
        self.che_do_sheet = "vi_tri"
        self.vi_tri_sheet = 1
        # Mặc định "1" (không phải "1,2") — để tránh trường hợp người dùng chọn
        # radio "Theo danh sách/khoảng sheet" nhưng quên sửa nội dung ô nhập, vô
        # tình in luôn nhiều sheet hơn mong muốn (giống lỗi đã sửa bên PDF).
        self.danh_sach_vi_tri_sheet = "1"
        self.ten_sheet = ""

        # Trang cần in trong mỗi sheet Excel, dạng danh sách/khoảng, vd "1,2" hoặc "3-4"
        # hoặc "1,3,5-7" (in trang 1, 3, và từ 5 đến 7).
        # Mặc định "1" (không phải "1,2") — cùng lý do như trên, tránh in nhầm
        # nhiều trang hơn mong muốn nếu quên sửa ô nhập.
        self.danh_sach_trang_excel = "1"

        # Chế độ trang áp dụng riêng cho file PDF (chỉ áp dụng khi loai_file == "pdf"):
        #   "toan_bo"   -> in nguyên file PDF gốc, không cắt trang
        #   "chi_dinh"  -> chỉ in đúng danh sách/khoảng trang riêng self.danh_sach_trang_pdf
        self.che_do_trang_pdf = "toan_bo"
        # Mặc định "1" (không phải "1,2") — để tránh trường hợp người dùng chọn
        # radio "Chỉ in trang chỉ định" nhưng quên sửa nội dung ô nhập, vô tình
        # in luôn nhiều trang hơn mong muốn thay vì đúng 1 trang họ định chọn.
        self.danh_sach_trang_pdf = "1"

        # Số bản in mặc định — áp dụng cho loại file đang được chọn (Excel hoặc PDF).
        self.so_ban_in = 1

        self._dang_in = False
        self._worker_thread = None
        self._dang_cho_dong_cua_so = False

    # ---------- Màn hình chọn thư mục ----------
    def _build_folder_select_ui(self):
        for w in self.root.winfo_children():
            w.destroy()

        frame = tk.Frame(self.root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text=TEN_PROJECT, font=("Segoe UI", 15, "bold"), fg="#1565c0").pack(pady=(0, 5))
        tk.Label(frame, text="Chọn thư mục cần quét", font=("Segoe UI", 13, "bold")).pack(pady=(0, 15))
        tk.Label(
            frame,
            text="Chương trình sẽ quét thư mục bạn chọn (bao gồm cả các thư mục con bên trong).\nBước tiếp theo bạn sẽ chọn quét file Excel hay file PDF.",
            fg="#666", justify="center", wraplength=460,
        ).pack(pady=(0, 20))

        duong_dan_var = tk.StringVar(value="(Chưa chọn thư mục)")
        duong_dan_label = tk.Label(frame, textvariable=duong_dan_var, wraplength=460, justify="center", fg="#2e7d32")
        duong_dan_label.pack(pady=(0, 20))

        tiep_tuc_btn = {"btn": None}

        def chon_thu_muc():
            duong_dan = filedialog.askdirectory(title="Chọn thư mục chứa file Excel cần in")
            if not duong_dan:
                return
            self.thu_muc_goc = Path(duong_dan)
            duong_dan_var.set(str(self.thu_muc_goc))
            if tiep_tuc_btn["btn"] is not None:
                tiep_tuc_btn["btn"].config(state="normal")

        tk.Button(
            frame, text="Chọn thư mục...", width=20, bg="#1565c0", fg="white",
            font=("Segoe UI", 10, "bold"), command=chon_thu_muc
        ).pack(pady=(0, 10))

        def tiep_tuc():
            if self.thu_muc_goc is None:
                messagebox.showwarning("Chưa chọn thư mục", "Vui lòng chọn thư mục trước khi tiếp tục.")
                return
            self._build_loai_file_ui()

        btn = tk.Button(
            frame, text="Tiếp tục", width=20, bg="#2e7d32", fg="white",
            font=("Segoe UI", 10, "bold"), command=tiep_tuc, state="disabled"
        )
        btn.pack(pady=10)
        tiep_tuc_btn["btn"] = btn

    def _build_loai_file_ui(self):
        for w in self.root.winfo_children():
            w.destroy()

        frame = tk.Frame(self.root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Chọn loại file cần in", font=("Segoe UI", 13, "bold")).pack(pady=(10, 5))
        tk.Label(frame, text=f"Thư mục: {self.thu_muc_goc}", fg="#666", wraplength=460).pack(pady=(0, 5))
        tk.Label(
            frame,
            text="Chỉ chọn được 1 loại — sau khi chọn, loại còn lại sẽ bị bỏ qua\nhoàn toàn (không quét, không hiện cài đặt liên quan).",
            fg="#666", justify="center", wraplength=460,
        ).pack(pady=(0, 25))

        def chon(loai):
            self.loai_file = loai
            self._quet_thu_muc()
            self._build_settings_ui()

        btn_frame = tk.Frame(frame)
        btn_frame.pack(pady=10)
        tk.Button(
            btn_frame, text="📊  File Excel\n(.xls / .xlsx)", width=18, height=4,
            bg="#2e7d32", fg="white", font=("Segoe UI", 11, "bold"),
            command=lambda: chon("excel")
        ).pack(side="left", padx=15)
        tk.Button(
            btn_frame, text="📄  File PDF\n(.pdf)", width=18, height=4,
            bg="#c62828", fg="white", font=("Segoe UI", 11, "bold"),
            command=lambda: chon("pdf")
        ).pack(side="left", padx=15)

        tk.Button(frame, text="Đổi thư mục", command=self._build_folder_select_ui).pack(pady=(25, 0))

    def _quet_thu_muc(self):
        duoi_hop_le = (".xls", ".xlsx") if self.loai_file == "excel" else (".pdf",)
        so_thay = set()
        danh_sach_file = []
        for p in self.thu_muc_goc.rglob("*"):
            if p.suffix.lower() not in duoi_hop_le or p.name.startswith("~$"):
                continue
            duong_dan_that = p.resolve()
            if duong_dan_that in so_thay:
                continue
            so_thay.add(duong_dan_that)
            danh_sach_file.append(p)
        # Sắp xếp "tự nhiên" theo đường dẫn tương đối để thứ tự in ổn định, dễ đoán,
        # và đúng thứ tự số (1, 2, ..., 10) thay vì so sánh chuỗi thuần (1, 10, 2).
        danh_sach_file.sort(key=lambda p: khoa_sap_xep_tu_nhien(str(p.relative_to(self.thu_muc_goc))))
        self.danh_sach_file = danh_sach_file

    def _khoi_dong(self):
        self._don_dep_file_tam_cu()

        self.root = tk.Tk()
        self.root.title(TEN_PROJECT)
        self.root.geometry("700x650")
        self.root.minsize(600, 500)
        self.root.resizable(True, True)
        self.root.protocol("WM_DELETE_WINDOW", self._xu_ly_dong_cua_so)

        self._build_folder_select_ui()
        self.root.mainloop()

    def _don_dep_file_tam_cu(self):
        """
        Xóa các file PDF tạm còn sót lại từ những lần chạy trước (ví dụ do chương
        trình bị tắt đột ngột trước khi kịp tự dọn). Chỉ xóa file đã tạo hơn 1 giờ
        để không đụng vào file tạm của một lượt in vừa mới chạy gần đây.
        """
        try:
            thu_muc_tam = Path(tempfile.gettempdir()) / "InHangLoat_ByAnDuc_temp"
            if not thu_muc_tam.exists():
                return
            bay_gio = time.time()
            for f in thu_muc_tam.glob("*.pdf"):
                try:
                    if bay_gio - f.stat().st_mtime > 3600:
                        f.unlink()
                except Exception:
                    pass  # file có thể đang được ứng dụng khác giữ, bỏ qua, không quan trọng
        except Exception:
            pass

    def _xu_ly_dong_cua_so(self):
        """Chặn việc đóng đột ngột khi đang in, tránh để tiến trình Excel chạy ngầm sót lại."""
        if self._dang_in:
            dong_y = messagebox.askyesno(
                "Đang in dở",
                "Chương trình đang gửi lệnh in.\n"
                "Nếu thoát ngay bây giờ, các lệnh in đã gửi vẫn nằm trong hàng đợi máy in,\n"
                "nhưng chương trình sẽ dừng gửi thêm lệnh mới.\n\n"
                "Bạn có chắc muốn thoát không?"
            )
            if not dong_y:
                return
            self.stop_event.set()
            # QUAN TRỌNG: không đóng cửa sổ ngay — luồng in chạy dạng daemon thread,
            # nếu tiến trình chính thoát trước khi nó kịp chạy xong excel.Quit(), Excel
            # sẽ bị giết đột ngột giữa chừng và để lại tiến trình EXCEL.EXE chạy ngầm.
            # Phải đợi luồng in thực sự kết thúc rồi mới đóng cửa sổ.
            self._dang_cho_dong_cua_so = True
            self._vo_hieu_hoa_cua_so_dang_dong()
            self._doi_worker_roi_dong()
            return
        self.root.destroy()

    def _vo_hieu_hoa_cua_so_dang_dong(self):
        for w in self.root.winfo_children():
            w.destroy()
        frame = tk.Frame(self.root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)
        tk.Label(
            frame, text="Đang dừng và dọn dẹp, vui lòng chờ...",
            font=("Segoe UI", 12, "bold")
        ).pack(expand=True)

    def _doi_worker_roi_dong(self):
        if self._worker_thread is not None and self._worker_thread.is_alive():
            self.root.after(200, self._doi_worker_roi_dong)
            return
        self.root.destroy()

    # ---------- Màn hình cài đặt: chọn sheet + khoảng trang ----------
    # ---------- Khung có thể cuộn dọc (dùng cho màn hình nhiều nội dung) ----------
    def _tao_khung_cuon(self, cha):
        """
        Tạo 1 khung có thể cuộn dọc bên trong widget `cha`, trả về khung con (Frame)
        để đặt nội dung vào — thay vì bị bóp méo/cắt cụt khi tổng chiều cao nội dung
        vượt quá chiều cao cửa sổ (Frame thường không tự cuộn được), giờ người dùng
        chỉ cần cuộn chuột hoặc kéo thanh cuộn để xem hết.
        """
        canvas = tk.Canvas(cha, highlightthickness=0)
        scrollbar = tk.Scrollbar(cha, orient="vertical", command=canvas.yview)
        khung_trong = tk.Frame(canvas, padx=20, pady=20)

        khung_trong.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        cua_so_canvas = canvas.create_window((0, 0), window=khung_trong, anchor="nw")
        # Khi cửa sổ được kéo giãn ngang, khung con cũng giãn theo đúng độ rộng canvas.
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(cua_so_canvas, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Chỉ bật cuộn bằng con lăn chuột khi con trỏ đang thực sự ở trong khung này,
        # và tắt đi khi rời khỏi — tránh ảnh hưởng tới các màn hình khác không cuộn.
        def _cuon_chuot(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _cuon_chuot))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        return khung_trong

    def _build_settings_ui(self):
        for w in self.root.winfo_children():
            w.destroy()

        frame = self._tao_khung_cuon(self.root)

        ten_loai = "Excel" if self.loai_file == "excel" else "PDF"
        tk.Label(frame, text=f"Cài đặt in — File {ten_loai}", font=("Segoe UI", 13, "bold")).pack(pady=(0, 5))
        tk.Label(frame, text=f"Thư mục: {self.thu_muc_goc}", fg="#666", wraplength=460).pack(pady=(0, 15))

        loi_label = tk.Label(frame, text="", fg="#c62828", wraplength=460, justify="left")

        # Các biến/hàm dưới đây được gán tùy theo self.loai_file — chỉ nhánh tương
        # ứng mới thực sự tạo widget, nhánh còn lại không hiện ra chút nào.
        che_do_var = danh_sach_var = vi_tri_var = ten_var = trang_excel_var = None
        che_do_pdf_var = trang_pdf_var = None

        if self.loai_file == "excel":
            # --- Chọn sheet ---
            tk.Label(frame, text="Sheet cần in:", font=("Segoe UI", 10, "bold"), anchor="w").pack(fill="x")

            che_do_var = tk.StringVar(value=self.che_do_sheet)
            vi_tri_var = tk.StringVar(value=str(self.vi_tri_sheet))
            danh_sach_var = tk.StringVar(value=self.danh_sach_vi_tri_sheet)
            ten_var = tk.StringVar(value=self.ten_sheet)

            row1 = tk.Frame(frame)
            row1.pack(fill="x", pady=(5, 0))
            tk.Radiobutton(row1, text="Theo số thứ tự sheet:", variable=che_do_var, value="vi_tri").pack(side="left")
            tk.Entry(row1, textvariable=vi_tri_var, width=5).pack(side="left", padx=5)
            tk.Label(row1, text="(1 = sheet đầu tiên)", fg="#666").pack(side="left")

            row1b = tk.Frame(frame)
            row1b.pack(fill="x", pady=(5, 0))
            tk.Radiobutton(row1b, text="Theo danh sách/khoảng sheet:", variable=che_do_var, value="danh_sach").pack(side="left")
            tk.Entry(row1b, textvariable=danh_sach_var, width=12).pack(side="left", padx=5)
            tk.Label(row1b, text="(vd: 1,2  hoặc  3-4  hoặc  1-2,4)", fg="#666").pack(side="left")

            row2 = tk.Frame(frame)
            row2.pack(fill="x", pady=(5, 0))
            tk.Radiobutton(row2, text="Theo tên sheet:", variable=che_do_var, value="ten").pack(side="left")
            tk.Entry(row2, textvariable=ten_var, width=15).pack(side="left", padx=5)
            tk.Label(row2, text="(áp dụng cho mọi file)", fg="#666").pack(side="left")

            row3 = tk.Frame(frame)
            row3.pack(fill="x", pady=(5, 10))
            tk.Radiobutton(row3, text="In tất cả các sheet trong file", variable=che_do_var, value="tat_ca").pack(side="left")

            # --- Chọn trang cần in ---
            tk.Label(frame, text="Trang cần in trong mỗi sheet:",
                     font=("Segoe UI", 10, "bold"), anchor="w").pack(fill="x", pady=(10, 0))

            trang_excel_var = tk.StringVar(value=self.danh_sach_trang_excel)

            row4 = tk.Frame(frame)
            row4.pack(fill="x", pady=(5, 0))
            tk.Entry(row4, textvariable=trang_excel_var, width=20).pack(side="left")
            tk.Label(row4, text="(vd: 1,2  hoặc  3-4  hoặc  1,3,5-7)", fg="#666").pack(side="left", padx=5)

            row4b = tk.Frame(frame)
            row4b.pack(fill="x", pady=(2, 0))
            tk.Label(
                row4b,
                text="⚠ Nhớ kiểm tra lại đúng nội dung trong các ô nhập ở trên (sheet lẫn\n"
                     "trang) trước khi bấm Tiếp tục — chương trình luôn in đúng theo nội\n"
                     "dung đang hiển thị trong ô, kể cả khi bạn quên sửa lại.",
                fg="#c62828", justify="left"
            ).pack(side="left")

        else:  # self.loai_file == "pdf"
            # --- Chọn chế độ trang ---
            tk.Label(frame, text="Trang cần in:", font=("Segoe UI", 10, "bold"), anchor="w").pack(fill="x")

            che_do_pdf_var = tk.StringVar(value=self.che_do_trang_pdf)

            row5 = tk.Frame(frame)
            row5.pack(fill="x", pady=(5, 0))
            tk.Radiobutton(row5, text="In toàn bộ trang trong file", variable=che_do_pdf_var, value="toan_bo").pack(side="left")

            row6 = tk.Frame(frame)
            row6.pack(fill="x", pady=(5, 0))
            tk.Radiobutton(row6, text="Chỉ in trang chỉ định:", variable=che_do_pdf_var, value="chi_dinh").pack(side="left")

            trang_pdf_var = tk.StringVar(value=self.danh_sach_trang_pdf)
            tk.Entry(row6, textvariable=trang_pdf_var, width=20).pack(side="left", padx=(10, 0))

            row6b = tk.Frame(frame)
            row6b.pack(fill="x", pady=(2, 0))
            tk.Label(row6b, text="(vd: 1,2  hoặc  3-4  hoặc  1,3,5-7)", fg="#666").pack(side="left", padx=(115, 0))

            row6c = tk.Frame(frame)
            row6c.pack(fill="x", pady=(2, 0))
            tk.Label(
                row6c,
                text="⚠ Nhớ kiểm tra lại đúng số trang trong ô trên trước khi bấm Tiếp tục —\n"
                     "nếu chọn \"Chỉ in trang chỉ định\" mà quên sửa ô nhập, chương trình sẽ\n"
                     "in đúng theo nội dung đang hiện trong ô đó.",
                fg="#c62828", justify="left"
            ).pack(side="left", padx=(115, 0))

        # --- Số bản in mặc định (áp dụng cho toàn bộ file thuộc loại đã chọn) ---
        tk.Label(frame, text="Số bản in:", font=("Segoe UI", 10, "bold"), anchor="w").pack(fill="x", pady=(15, 0))

        so_ban_var = tk.StringVar(value=str(self.so_ban_in))

        row7 = tk.Frame(frame)
        row7.pack(fill="x", pady=(5, 0))
        tk.Entry(row7, textvariable=so_ban_var, width=6).pack(side="left")
        tk.Label(row7, text="bản mỗi file", fg="#666").pack(side="left", padx=(8, 0))

        loi_label.pack(pady=(10, 0))

        def xac_nhan_cai_dat():
            loi_label.config(text="")

            if self.loai_file == "excel":
                che_do = che_do_var.get()
                if che_do == "vi_tri":
                    try:
                        vi_tri = int(vi_tri_var.get())
                        if vi_tri < 1:
                            raise ValueError
                    except ValueError:
                        loi_label.config(text="Số thứ tự sheet phải là số nguyên dương (VD: 1, 2, 3...).")
                        return
                    self.vi_tri_sheet = vi_tri

                elif che_do == "danh_sach":
                    try:
                        parse_danh_sach_so(danh_sach_var.get())
                    except ValueError as e:
                        loi_label.config(text=f"Danh sách sheet không hợp lệ: {e}")
                        return
                    self.danh_sach_vi_tri_sheet = danh_sach_var.get().strip()

                elif che_do == "ten":
                    ten = ten_var.get().strip()
                    if not ten:
                        loi_label.config(text="Vui lòng nhập tên sheet.")
                        return
                    self.ten_sheet = ten

                try:
                    parse_danh_sach_so(trang_excel_var.get())
                except ValueError as e:
                    loi_label.config(text=f"Trang cần in không hợp lệ: {e}")
                    return

                self.che_do_sheet = che_do
                self.danh_sach_trang_excel = trang_excel_var.get().strip()

            else:  # pdf
                che_do_pdf = che_do_pdf_var.get()
                if che_do_pdf == "chi_dinh":
                    try:
                        parse_danh_sach_so(trang_pdf_var.get())
                    except ValueError as e:
                        loi_label.config(text=f"Trang cần in không hợp lệ: {e}")
                        return

                self.che_do_trang_pdf = che_do_pdf
                self.danh_sach_trang_pdf = trang_pdf_var.get().strip()

            try:
                so_ban = int(so_ban_var.get())
                if so_ban < 1:
                    raise ValueError
            except ValueError:
                loi_label.config(text="Số bản in phải là số nguyên dương (VD: 1, 2, 3...).")
                return
            self.so_ban_in = so_ban

            self._build_confirm_ui()

        btn_row = tk.Frame(frame)
        btn_row.pack(pady=20)
        tk.Button(
            btn_row, text="Đổi loại file", width=15,
            command=self._build_loai_file_ui
        ).pack(side="left", padx=10)
        tk.Button(
            btn_row, text="Tiếp tục", width=15, bg="#2e7d32", fg="white",
            font=("Segoe UI", 10, "bold"), command=xac_nhan_cai_dat
        ).pack(side="left", padx=10)

    # ---------- Màn hình xác nhận trước khi in ----------
    def _build_confirm_ui(self):
        for w in self.root.winfo_children():
            w.destroy()

        frame = tk.Frame(self.root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        so_file = len(self.danh_sach_file)

        if self.loai_file == "excel":
            if self.che_do_sheet == "vi_tri":
                mo_ta_sheet = f"sheet số {self.vi_tri_sheet}"
            elif self.che_do_sheet == "danh_sach":
                mo_ta_sheet = f"các sheet số {self.danh_sach_vi_tri_sheet}"
            elif self.che_do_sheet == "ten":
                mo_ta_sheet = f"sheet tên \"{self.ten_sheet}\""
            else:
                mo_ta_sheet = "tất cả các sheet"

            mo_ta_tom_tat = (
                f"Sẽ gửi lệnh in {so_file} file Excel\n"
                f"In {mo_ta_sheet}, trang {self.danh_sach_trang_excel} mỗi sheet, {self.so_ban_in} bản mỗi file"
            )
        else:
            mo_ta_pdf = "in toàn bộ trang" if self.che_do_trang_pdf == "toan_bo" else f"chỉ in trang {self.danh_sach_trang_pdf}"

            mo_ta_tom_tat = (
                f"Sẽ gửi lệnh in {so_file} file PDF\n"
                f"{mo_ta_pdf}, {self.so_ban_in} bản mỗi file"
            )

        tk.Label(
            frame, text=mo_ta_tom_tat, font=("Segoe UI", 12, "bold"), justify="center",
        ).pack(pady=(0, 10))

        if so_file == 0:
            ten_loai = "Excel" if self.loai_file == "excel" else "PDF"
            tk.Label(frame, text=f"Không tìm thấy file {ten_loai} nào trong thư mục này.", fg="#c62828").pack()
            tk.Button(frame, text="Quay lại", width=15, command=self._build_settings_ui).pack(pady=15)
            return

        tk.Label(
            frame, text="Mẹo: nhấn giữ chuột và kéo thả 1 dòng để sắp xếp lại thứ tự in.",
            fg="#666"
        ).pack(pady=(0, 5))

        list_frame = tk.Frame(frame)
        list_frame.pack(fill="both", expand=True)

        scrollbar_y = tk.Scrollbar(list_frame, orient="vertical")
        scrollbar_y.pack(side="right", fill="y")

        # scrollbar_x phải nằm CHUNG trong list_frame (không phải frame gốc), để nó bị
        # giới hạn đúng trong phạm vi khung danh sách — pack trước listbox với side="bottom"
        # để nó bám sát ngay dưới danh sách, thay vì bị đẩy xuống đáy toàn bộ cửa sổ
        # (dưới cả hàng nút bấm) như khi gắn vào frame gốc.
        scrollbar_x = tk.Scrollbar(list_frame, orient="horizontal")
        scrollbar_x.pack(side="bottom", fill="x")

        listbox = tk.Listbox(
            list_frame, yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set,
            height=14
        )

        def lam_moi_danh_sach_hien_thi():
            listbox.delete(0, "end")
            for idx, p in enumerate(self.danh_sach_file, start=1):
                try:
                    duong_dan_hien_thi = str(p.relative_to(self.thu_muc_goc))
                except ValueError:
                    duong_dan_hien_thi = str(p)
                listbox.insert("end", f"{idx}. {duong_dan_hien_thi}")

        # ----- Kéo thả để sắp xếp lại thứ tự -----
        trang_thai_keo = {"chi_so": None}

        def bat_dau_keo(event):
            trang_thai_keo["chi_so"] = listbox.nearest(event.y)

        def dang_keo(event):
            chi_so_cu = trang_thai_keo["chi_so"]
            if chi_so_cu is None:
                return
            chi_so_moi = listbox.nearest(event.y)
            if chi_so_moi == chi_so_cu or not (0 <= chi_so_moi < len(self.danh_sach_file)):
                return
            self.danh_sach_file[chi_so_cu], self.danh_sach_file[chi_so_moi] = (
                self.danh_sach_file[chi_so_moi], self.danh_sach_file[chi_so_cu]
            )
            lam_moi_danh_sach_hien_thi()
            listbox.selection_clear(0, "end")
            listbox.selection_set(chi_so_moi)
            trang_thai_keo["chi_so"] = chi_so_moi

        def ket_thuc_keo(event):
            trang_thai_keo["chi_so"] = None

        listbox.bind("<Button-1>", bat_dau_keo)
        listbox.bind("<B1-Motion>", dang_keo)
        listbox.bind("<ButtonRelease-1>", ket_thuc_keo)

        lam_moi_danh_sach_hien_thi()
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar_y.config(command=listbox.yview)
        scrollbar_x.config(command=listbox.xview)

        btn_frame = tk.Frame(frame)
        btn_frame.pack(pady=15)
        tk.Button(
            btn_frame, text="Bắt đầu in", width=15, bg="#2e7d32", fg="white",
            font=("Segoe UI", 10, "bold"), command=self._start_printing
        ).pack(side="left", padx=10)
        tk.Button(
            btn_frame, text="Quay lại", width=15,
            command=self._build_settings_ui
        ).pack(side="left", padx=10)
        tk.Button(
            btn_frame, text="Hủy bỏ", width=15,
            command=self.root.destroy
        ).pack(side="left", padx=10)

    # ---------- Màn hình tiến trình khi đang in ----------
    def _build_progress_ui(self):
        for w in self.root.winfo_children():
            w.destroy()

        frame = tk.Frame(self.root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Đang gửi lệnh in...", font=("Segoe UI", 12, "bold")).pack(pady=(0, 10))

        self.progress = ttk.Progressbar(
            frame, orient="horizontal", length=460, mode="determinate",
            maximum=len(self.danh_sach_file)
        )
        self.progress.pack(pady=(0, 10))

        self.status_label = tk.Label(frame, text="", wraplength=460, justify="left")
        self.status_label.pack(pady=(0, 10))

        log_frame = tk.Frame(frame)
        log_frame.pack(fill="both", expand=True)
        scrollbar = tk.Scrollbar(log_frame)
        scrollbar.pack(side="right", fill="y")
        self.log_box = tk.Listbox(log_frame, yscrollcommand=scrollbar.set, height=10)
        self.log_box.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.log_box.yview)

        self.cancel_btn = tk.Button(
            frame, text="HỦY TẤT CẢ LỆNH IN", bg="#c62828", fg="white",
            font=("Segoe UI", 10, "bold"), command=self._cancel_printing
        )
        self.cancel_btn.pack(pady=15)

    def _start_printing(self):
        self.stop_event.clear()
        self._dang_in = True
        self._build_progress_ui()
        self._worker_thread = threading.Thread(target=self._print_worker, daemon=True)
        self._worker_thread.start()
        self.root.after(100, self._poll_queue)

    def _cancel_printing(self):
        self.stop_event.set()
        self.cancel_btn.config(state="disabled", text="Đang hủy, chờ file hiện tại in xong...")

    def _poll_queue(self):
        if self._dang_cho_dong_cua_so:
            return  # Cửa sổ đang được đóng theo yêu cầu người dùng, không cập nhật widget nữa.
        try:
            while True:
                msg = self.log_queue.get_nowait()
                if msg[0] == "progress":
                    _, idx, ten_file, trang_thai = msg
                    self.progress["value"] = idx
                    self.status_label.config(text=f"[{idx}/{len(self.danh_sach_file)}] {ten_file}")
                    self.log_box.insert("end", f"{ten_file}  ->  {trang_thai}")
                    # Giới hạn tối đa 500 dòng hiển thị trong log — nếu quét hàng nghìn
                    # file, Listbox phình quá to sẽ làm giao diện phản hồi chậm dần.
                    # Xóa bớt dòng cũ nhất khi vượt ngưỡng (thông tin đầy đủ vẫn có
                    # trong màn tổng kết cuối cùng, chỉ log chi tiết bị rút gọn).
                    if self.log_box.size() > 500:
                        self.log_box.delete(0)
                    self.log_box.see("end")
                elif msg[0] == "done":
                    _, thanh_cong, loi, huy = msg
                    self._dang_in = False
                    self._show_summary(thanh_cong, loi, huy)
                    return
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _show_summary(self, thanh_cong, loi, huy):
        for w in self.root.winfo_children():
            w.destroy()
        frame = tk.Frame(self.root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        tieu_de = "Đã hủy" if huy else "Hoàn thành"
        mau = "#c62828" if huy else "#2e7d32"
        tk.Label(frame, text=tieu_de, font=("Segoe UI", 14, "bold"), fg=mau).pack(pady=(0, 15))
        tk.Label(frame, text=f"Lệnh in thành công: {thanh_cong}\nLệnh in bị lỗi: {loi}", font=("Segoe UI", 11)).pack()

        if huy:
            tk.Label(
                frame,
                text=("Lưu ý: những file đã gửi lệnh in TRƯỚC KHI bấm hủy\n"
                      "vẫn có thể đang nằm trong hàng đợi máy in (Print Queue)\n"
                      "vì lệnh in đã được gửi đi rồi. Hãy mở mục quản lý máy in\n"
                      "trong Windows để xóa các lệnh in đó nếu cần."),
                fg="#c62828", wraplength=460, justify="left"
            ).pack(pady=15)

        co_file_pdf = any(p.suffix.lower() == ".pdf" for p in self.danh_sach_file)
        if co_file_pdf:
            tk.Label(
                frame,
                text=("Lưu ý về file PDF: chương trình gửi lệnh in qua trình đọc PDF\n"
                      "mặc định của máy (Edge/Acrobat...). Nếu trình đọc đó gặp lỗi nội bộ\n"
                      "khi đang in, chương trình sẽ không phát hiện được (do đây là giới hạn\n"
                      "của cơ chế in gián tiếp qua Windows). Vui lòng kiểm tra thực tế bản in."),
                fg="#666", wraplength=460, justify="left"
            ).pack(pady=(10, 0))

        btn_row = tk.Frame(frame)
        btn_row.pack(pady=10)
        tk.Button(
            btn_row, text="Quét thư mục khác", width=18,
            command=self._build_folder_select_ui
        ).pack(side="left", padx=5)
        tk.Button(btn_row, text="Đóng", width=15, command=self.root.destroy).pack(side="left", padx=5)

    # ---------- In riêng cho file PDF (không cần Excel) ----------
    def _gui_lenh_in_pdf(self, duong_dan_pdf: Path, so_ban_in: int, la_ngang: bool) -> bool:
        """
        Gửi ĐÚNG 1 lệnh in vật lý cho file PDF, với số bản in VÀ hướng giấy được
        đặt CHÍNH XÁC qua driver máy in mặc định — KHÔNG lặp lại lệnh in nhiều
        lần như cách cũ.

        LÝ DO: verb "print" mặc định của Windows không có tham số riêng để chọn
        số bản in hay hướng giấy — nó luôn in theo đúng số bản/hướng giấy ĐANG
        ĐƯỢC LƯU SẴN trong driver máy in mặc định. Nếu driver đang để hướng dọc
        mà nội dung PDF thực tế nằm ngang, kết quả in ra sẽ bị co nhỏ lệch sang
        1 góc tờ giấy, trông như file "không in được". Cách đúng là đặt thẳng
        cả 2 thông số vào driver rồi chỉ gửi lệnh MỘT LẦN, để chính driver/máy
        in xử lý đúng — giống hệt cách Excel đã làm đúng từ trước qua tham số
        Copies riêng của nó (không phụ thuộc driver nên không dính lỗi này).

        QUAN TRỌNG: os.startfile(..., "print") là lệnh KHÔNG đồng bộ — nó chỉ
        yêu cầu Windows mở ứng dụng đọc PDF mặc định rồi trả về NGAY LẬP TỨC,
        không đợi ứng dụng đó thật sự đọc xong file và nộp lệnh in vào hàng đợi
        máy in. Nếu code xử lý file tiếp theo diễn ra quá sớm, trong khi file
        này còn đang mở dở, lệnh đó có thể "đè" lên trước khi ứng dụng đọc PDF
        kịp áp dụng đúng cấu hình — đây chính là nguyên nhân từng gây lỗi "file
        sau chỉ in 1 bản thay vì N bản". Vì vậy sau khi gửi lệnh, hàm này ĐỢI
        XÁC NHẬN THẬT SỰ (không chỉ đợi 1 khoảng thời gian cố định) bằng cách
        theo dõi hàng đợi máy in cho đến khi thấy lệnh in mới xuất hiện.

        Trả về True nếu xác nhận được lệnh in đã vào hàng đợi máy in, False
        nếu không xác nhận được trong thời gian chờ (không hẳn là thất bại,
        chỉ là không chắc chắn — xem thêm giới hạn trong
        doi_lenh_in_duoc_gui_vao_hang_doi()).
        """
        dat_cau_hinh_may_in_mac_dinh(so_ban_in, la_ngang)
        # Đợi 1 nhịp ngắn để chắc chắn Windows đã ghi nhận các thay đổi cấu hình
        # trước khi gửi lệnh, tránh trường hợp trình đọc PDF đọc phải giá trị cũ
        # do đọc quá sớm.
        time.sleep(0.3)

        danh_sach_id_truoc = _lay_danh_sach_job_id_dang_cho()
        os.startfile(str(duong_dan_pdf), "print")
        return doi_lenh_in_duoc_gui_vao_hang_doi(danh_sach_id_truoc)

    def _in_pdf(self, duong_dan_file: Path, so_ban_in: int):
        """
        Nếu self.che_do_trang_pdf == "toan_bo": gửi thẳng file PDF gốc đi in,
        không cần tạo file tạm (nhanh hơn, không phát sinh rác).
        Nếu == "chi_dinh": cắt đúng các trang trong self.danh_sach_trang_pdf (có thể là
        trang lẻ không liên tục, vd "3,5,7") ra 1 file PDF tạm rồi gửi file tạm đó đi in.
        Số bản in được xử lý bằng cách đặt thẳng vào driver máy in (xem
        _gui_lenh_in_pdf), KHÔNG lặp lại lệnh in và KHÔNG nhân bản nội dung
        trong file PDF — chỉ gửi đúng 1 lệnh in cho mỗi file.
        Trả về (mo_ta_trang_thai, so_thanh_cong, so_loi).
        """
        if self.che_do_trang_pdf == "toan_bo":
            try:
                # Xác định hướng giấy dựa vào trang ĐẦU TIÊN của file (đa số file
                # PDF thực tế đồng nhất hướng cho mọi trang; trường hợp file trộn
                # lẫn cả trang dọc và ngang là hiếm và nằm ngoài phạm vi xử lý ở đây).
                try:
                    reader_kiem_tra = PdfReader(str(duong_dan_file))
                    la_ngang = bool(reader_kiem_tra.pages) and kiem_tra_pdf_huong_ngang(reader_kiem_tra.pages[0])
                except Exception:
                    la_ngang = False

                da_xac_nhan = self._gui_lenh_in_pdf(duong_dan_file, so_ban_in, la_ngang)
                if da_xac_nhan:
                    return (f"Đã gửi lệnh in toàn bộ file (x{so_ban_in} bản)", 1, 0)
                return (
                    f"Đã gửi lệnh in toàn bộ file (x{so_ban_in} bản) — "
                    f"nhưng không xác nhận được lệnh đã vào hàng đợi máy in, vui lòng kiểm tra lại",
                    1, 0
                )
            except Exception as e:
                return (f"Lỗi khi in PDF: {e}", 0, 1)

        # ----- Chế độ chỉ in trang chỉ định -----
        try:
            danh_sach_trang = parse_danh_sach_so(self.danh_sach_trang_pdf, "Danh sách trang PDF")

            reader = PdfReader(str(duong_dan_file))
            tong_trang = len(reader.pages)

            trang_hop_le = [t for t in danh_sach_trang if t <= tong_trang]
            trang_vuot_qua = [t for t in danh_sach_trang if t > tong_trang]

            if not trang_hop_le:
                return (f"Lỗi: file chỉ có {tong_trang} trang, không có trang nào trong {danh_sach_trang} hợp lệ", 0, 1)

            writer = PdfWriter()
            for so_trang in trang_hop_le:
                writer.add_page(reader.pages[so_trang - 1])

            # Xác định hướng giấy dựa vào trang ĐẦU TIÊN trong danh sách trang
            # thực sự được in ra (không phải trang đầu file, mà đúng trang được
            # người dùng chọn).
            la_ngang = kiem_tra_pdf_huong_ngang(reader.pages[trang_hop_le[0] - 1])

            thu_muc_tam = Path(tempfile.gettempdir()) / "InHangLoat_ByAnDuc_temp"
            thu_muc_tam.mkdir(exist_ok=True)
            # Không nhét danh sách số trang vào tên file tạm — nếu người dùng nhập nhiều
            # trang lẻ (vd "1,3,5,7,9,11,...") cộng với tên file gốc vốn đã dài, đường dẫn
            # rất dễ vượt quá giới hạn 260 ký tự (MAX_PATH) của Windows, gây lỗi
            # "File name too long". uuid phía sau đã đủ đảm bảo không trùng tên rồi.
            # Đồng thời cắt bớt phần tên gốc (chỉ giữ 60 ký tự đầu) để an toàn hơn nữa
            # với những file có tên cực dài.
            ten_goc_rut_gon = duong_dan_file.stem[:60]
            file_tam = thu_muc_tam / f"{ten_goc_rut_gon}_trichtrang_{uuid.uuid4().hex[:8]}.pdf"

            with open(file_tam, "wb") as f:
                writer.write(f)

            da_xac_nhan = self._gui_lenh_in_pdf(file_tam, so_ban_in, la_ngang)

            # Lệnh in ở trên là KHÔNG đồng bộ. Sau khi đã đợi xác nhận lệnh in vào
            # hàng đợi máy in (xem _gui_lenh_in_pdf), file tạm vẫn có thể đang được
            # trình đọc PDF giữ trong lúc render/in — không xóa ngay, mà hẹn giờ xóa
            # sau 2 phút — đủ thời gian để hầu hết các trình đọc PDF đọc xong file.
            # Nếu lần xóa đầu thất bại (file vẫn đang bị trình đọc PDF giữ), thử lại
            # thêm 2 lần nữa cách nhau 2 phút; nếu vẫn thất bại thì để lần khởi động
            # chương trình tiếp theo tự dọn rác cũ (>1 giờ) xử lý nốt.
            def _xoa_file_tam(duong_dan=file_tam, so_lan_con_lai=3):
                try:
                    if duong_dan.exists():
                        duong_dan.unlink()
                    return
                except Exception:
                    pass  # file có thể vẫn đang được trình đọc PDF giữ
                if so_lan_con_lai > 1:
                    hen_gio_lai = threading.Timer(
                        120, _xoa_file_tam, kwargs={"duong_dan": duong_dan, "so_lan_con_lai": so_lan_con_lai - 1}
                    )
                    hen_gio_lai.daemon = True
                    hen_gio_lai.start()

            hen_gio = threading.Timer(120, _xoa_file_tam)
            hen_gio.daemon = True
            hen_gio.start()

            mo_ta = f"Đã gửi lệnh in trang {', '.join(str(t) for t in trang_hop_le)} (x{so_ban_in} bản)"
            if trang_vuot_qua:
                mo_ta += f" (bỏ qua trang {', '.join(str(t) for t in trang_vuot_qua)} vì file chỉ có {tong_trang} trang)"
            if not da_xac_nhan:
                mo_ta += " — không xác nhận được lệnh đã vào hàng đợi máy in, vui lòng kiểm tra lại"

            return (mo_ta, 1, 0)

        except Exception as e:
            return (f"Lỗi khi in PDF: {e}", 0, 1)

    # ---------- Luồng xử lý in, chạy nền để không đứng giao diện ----------
    def _print_worker(self):
        pythoncom.CoInitialize()
        excel = None
        so_sheet_thanh_cong = 0
        so_sheet_loi = 0
        bi_huy = False
        loi_mo_excel = None

        can_excel = (self.loai_file == "excel")
        khoang_trang_excel = []
        loi_trang_excel = None
        if can_excel:
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                excel.ScreenUpdating = False
            except Exception as e:
                # Không mở được Excel thì các file Excel sẽ báo lỗi riêng từng file.
                loi_mo_excel = str(e)

            try:
                # Gom các trang lẻ liền kề (vd 1,2,3) thành khoảng (1,3) để giảm số lệnh
                # in thực tế phải gửi cho Excel — Excel chỉ nhận From/To liên tục mỗi lệnh.
                khoang_trang_excel = gom_nhom_lien_tuc(
                    parse_danh_sach_so(self.danh_sach_trang_excel, "Trang Excel")
                )
            except ValueError as e:
                loi_trang_excel = str(e)

        for idx, duong_dan_file in enumerate(self.danh_sach_file, start=1):
            # Kiểm tra cờ hủy TRƯỚC khi xử lý file tiếp theo.
            # Lưu ý: nếu file hiện tại đang trong lúc gửi lệnh in thì lệnh đó
            # vẫn hoàn tất bình thường, chỉ các file CHƯA xử lý mới thực sự bị dừng lại.
            if self.stop_event.is_set():
                bi_huy = True
                break

            # ----- File PDF -----
            if self.loai_file == "pdf":
                mo_ta, thanh_cong, loi = self._in_pdf(duong_dan_file, self.so_ban_in)
                so_sheet_thanh_cong += thanh_cong
                so_sheet_loi += loi
                self.log_queue.put(("progress", idx, duong_dan_file.name, mo_ta))
                continue

            # ----- File Excel -----
            if excel is None:
                so_sheet_loi += 1
                self.log_queue.put((
                    "progress", idx, duong_dan_file.name,
                    f"Lỗi: không mở được Excel ({loi_mo_excel})"
                ))
                continue

            if loi_trang_excel is not None:
                so_sheet_loi += 1
                self.log_queue.put((
                    "progress", idx, duong_dan_file.name,
                    f"Lỗi: trang Excel không hợp lệ ({loi_trang_excel})"
                ))
                continue

            wb = None
            ket_qua_sheets = []
            try:
                wb = excel.Workbooks.Open(str(duong_dan_file), ReadOnly=True)

                # Xác định danh sách sheet cần in cho file này theo cấu hình đã chọn.
                if self.che_do_sheet == "vi_tri":
                    danh_sach_ws = [("vi_tri", self.vi_tri_sheet)]
                elif self.che_do_sheet == "danh_sach":
                    danh_sach_ws = [("vi_tri", vt) for vt in parse_danh_sach_so(self.danh_sach_vi_tri_sheet)]
                elif self.che_do_sheet == "ten":
                    danh_sach_ws = [("ten", self.ten_sheet)]
                else:  # tat_ca
                    danh_sach_ws = [("obj", ws) for ws in wb.Sheets]

                for loai, tham_chieu in danh_sach_ws:
                    ws = None
                    try:
                        # "vi_tri" và "ten" đều gọi wb.Sheets(...) giống nhau,
                        # chỉ khác kiểu tham chiếu (số thứ tự hoặc tên).
                        ws = tham_chieu if loai == "obj" else wb.Sheets(tham_chieu)

                        # Mỗi khoảng liên tục trong khoang_trang_excel là 1 lệnh in riêng,
                        # vì Excel PrintOut chỉ nhận From/To liên tục cho mỗi lệnh.
                        # Copies=self.so_ban_in: Excel tự in đúng số bản trong 1 lệnh,
                        # không cần lặp lại lệnh in nhiều lần.
                        for bat_dau, ket_thuc in khoang_trang_excel:
                            ws.PrintOut(From=bat_dau, To=ket_thuc, Copies=self.so_ban_in)

                        so_sheet_thanh_cong += 1
                        trang_da_in = ",".join(
                            f"{a}" if a == b else f"{a}-{b}" for a, b in khoang_trang_excel
                        )
                        ket_qua_sheets.append(f"{ws.Name}: đã in trang {trang_da_in} (x{self.so_ban_in} bản)")
                    except Exception as e:
                        so_sheet_loi += 1
                        ten_hien_thi = tham_chieu if loai != "obj" else "?"
                        ket_qua_sheets.append(f"Sheet {ten_hien_thi}: lỗi - {e}")
                    finally:
                        # Giải phóng tường minh ngay sau mỗi sheet, thay vì chỉ trông chờ
                        # vào việc biến ws bị gán lại ở vòng lặp kế tiếp.
                        ws = None

            except Exception as e:
                so_sheet_loi += 1
                ket_qua_sheets.append(f"Không mở được file: {e}")
            finally:
                danh_sach_ws = None  # giải phóng tham chiếu COM tới các worksheet (nếu có)
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                # Giải phóng rõ ràng tham chiếu COM (workbook, worksheet) thay vì chỉ
                # trông chờ vào garbage collector tự dọn — đây là nguyên nhân phổ biến
                # khiến EXCEL.EXE bị treo ngầm sau khi Quit() dù đã gọi.
                wb = None

            trang_thai = "; ".join(ket_qua_sheets) if ket_qua_sheets else "Không có sheet nào được in"
            self.log_queue.put(("progress", idx, duong_dan_file.name, trang_thai))
            # Gọi gc.collect() định kỳ (mỗi 5 file) thay vì sau MỖI file — vẫn đủ để
            # giải phóng COM object kịp thời, tránh Excel treo ngầm, mà đỡ tốn CPU hơn
            # khi in số lượng lớn file.
            if idx % 5 == 0:
                gc.collect()

        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            excel = None
            gc.collect()
        pythoncom.CoUninitialize()

        # Đặt lại số bản in và hướng giấy của driver máy in về mặc định (1 bản,
        # hướng dọc) sau khi in xong toàn bộ, để không làm ảnh hưởng tới các lần
        # in tay khác của người dùng sau này trên máy đó.
        dat_cau_hinh_may_in_mac_dinh(1, False)

        self.log_queue.put(("done", so_sheet_thanh_cong, so_sheet_loi, bi_huy))


if __name__ == "__main__":
    app = InPageApp()
    app._khoi_dong()

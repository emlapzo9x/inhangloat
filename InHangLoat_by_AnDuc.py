"""
In Hàng Loạt by An Duc v3.4 (Phiên bản PyQt6)
Công cụ in hàng loạt file Excel (.xls/.xlsx) và PDF (.pdf) trong 1 thư mục.
"""

import gc
import re
import tempfile
import threading
import queue
import time
from pathlib import Path

import win32com.client
import win32print
import win32gui
import win32ui
import win32con
import pythoncom
import pymupdf as fitz  # PyMuPDF — tên package mới; giữ alias "fitz" để không phải đổi code bên dưới
from PIL import Image, ImageWin

import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QPushButton, QFileDialog,
                             QMessageBox, QRadioButton, QLineEdit, QScrollArea,
                             QButtonGroup, QListWidget, QListWidgetItem, QProgressBar,
                             QAbstractItemView)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QCloseEvent


TEN_PROJECT = "In Hàng Loạt by An Duc v3.4"

# Hằng số hướng giấy của Windows (winspool.h): 1 = dọc (Portrait), 2 = ngang (Landscape).
DMORIENT_PORTRAIT = 1
DMORIENT_LANDSCAPE = 2

# =============================================================================
# LOGIC CORE (Được giữ nguyên hoàn toàn so với phiên bản Tkinter)
# =============================================================================

# =============================================================================
# IN PDF TRỰC TIẾP QUA GDI CỦA WINDOWS — KHÔNG THÔNG QUA BẤT KỲ ỨNG DỤNG ĐỌC
# PDF NÀO (Edge/Acrobat/Foxit...). Đây là thay đổi quan trọng so với các bản
# trước: bản cũ gọi verb "print" của Windows để nhờ ứng dụng đọc PDF mặc định
# tự mở file rồi tự in — cách đó có 2 rủi ro cố hữu không kiểm soát được:
#   1. "Bắn rồi quên" (fire-and-forget): nếu app đang bận xử lý file trước
#      (đặc biệt Edge/Chrome dùng chung 1 tiến trình cho mọi cửa sổ), yêu cầu
#      in file tiếp theo có thể bị bỏ qua âm thầm — gây MISS CẢ FILE.
#   2. App tự nhớ cấu hình (khoảng trang, số bản) từ lần in trước đó, có thể
#      âm thầm áp sai cho lần in im lặng tiếp theo — gây MISS TRANG.
# Cách làm mới này tự dựng cấu hình máy in (DEVMODE) HOÀN TOÀN RIÊNG cho mỗi
# lần in, tự render từng trang PDF thành ảnh (qua PyMuPDF) rồi gửi thẳng cho
# driver máy in qua GDI — toàn bộ quá trình đồng bộ (StartDoc/EndDoc chỉ trả
# về khi đã xử lý xong), không có khái niệm "bắn rồi quên", và không đọc/ghi
# gì vào cấu hình mặc định của hệ thống nên không có "trí nhớ" nào để bị dính.
#
# LƯU Ý VỀ TRIỂN KHAI: dùng win32gui/win32ui của pywin32 để tạo Device Context
# tùy chỉnh (chấp nhận truyền vào 1 DEVMODE riêng). 2 module này cần bộ
# Microsoft Visual C++ Redistributable cài sẵn trên máy chạy .exe — nếu gặp
# lỗi "DLL load failed while importing win32ui" khi mở app, cài bộ đó tại
# https://aka.ms/vs/17/release/vc_redist.x64.exe (miễn phí, chính chủ
# Microsoft) là đủ, không cần build lại gì cả.
# =============================================================================

def kiem_tra_pdf_huong_ngang(trang) -> bool:
    """
    Kiểm tra 1 trang PDF (đối tượng Page của PyMuPDF/fitz) có phải hướng ngang
    (landscape) hay không, dựa vào kích thước thực tế sau khi đã áp góc xoay
    (page.rect đã tự tính theo /Rotate của trang, không cần tự trừ hao thủ công
    như khi dùng pypdf).
    Trả về False (coi như hướng dọc) nếu không xác định được vì bất kỳ lý do gì.
    """
    try:
        hcn = trang.rect
        return hcn.width > hcn.height
    except Exception:
        return False


def in_truc_tiep_qua_gdi(tai_lieu, danh_sach_trang: list, so_ban_in: int, la_ngang: bool):
    """
    In trực tiếp danh sách trang (số thứ tự, bắt đầu từ 1) của 1 tài liệu
    PyMuPDF (fitz.Document) đã mở sẵn, tới máy in mặc định của Windows, qua
    GDI — không thông qua bất kỳ ứng dụng đọc PDF nào.

    Tự dựng 1 DEVMODE (cấu hình máy in: số bản + hướng giấy) HOÀN TOÀN RIÊNG
    cho lần in này qua win32gui.CreateDC(), KHÔNG đọc/ghi vào cấu hình mặc
    định của hệ thống (khác hẳn cách dùng win32print.SetPrinter) — nên không
    có rủi ro "dính" cấu hình từ lần in trước, và cũng không làm ảnh hưởng
    tới các lần in khác (kể cả in tay) sau này trên cùng máy in.

    Trả về None nếu thành công, hoặc chuỗi mô tả lỗi nếu thất bại.
    """
    try:
        ten_may_in = win32print.GetDefaultPrinter()
    except Exception as e:
        return f"Không tìm thấy máy in mặc định ({e})"

    hprinter = None
    hDC = None
    try:
        hprinter = win32print.OpenPrinter(ten_may_in)
        thong_tin = win32print.GetPrinter(hprinter, 2)
        devmode = thong_tin.get("pDevMode")
        if devmode is None:
            return "Không đọc được cấu hình máy in."
        devmode.Copies = max(1, so_ban_in)
        devmode.Orientation = DMORIENT_LANDSCAPE if la_ngang else DMORIENT_PORTRAIT

        # Dùng win32gui.CreateDC (không phải win32ui.CreateDC().CreatePrinterDC())
        # vì đây là cách DUY NHẤT trong pywin32 cho phép truyền vào 1 DEVMODE tùy
        # chỉnh khi tạo Device Context — bản CreatePrinterDC() đơn giản hơn không
        # nhận DEVMODE, sẽ luôn dùng lại cấu hình mặc định hệ thống.
        hdc_tho = win32gui.CreateDC("WINSPOOL", ten_may_in, devmode)
        hDC = win32ui.CreateDCFromHandle(hdc_tho)

        # Giới hạn độ phân giải render tối đa 200 DPI — đủ nét cho tài liệu văn
        # phòng thông thường, tránh ảnh quá nặng/chậm nếu máy in báo DPI rất cao.
        do_phan_giai = min(hDC.GetDeviceCaps(win32con.LOGPIXELSX), 200)
        ty_le_render = do_phan_giai / 72.0
        ma_tran = fitz.Matrix(ty_le_render, ty_le_render)

        rong_in = hDC.GetDeviceCaps(win32con.HORZRES)
        cao_in = hDC.GetDeviceCaps(win32con.VERTRES)

        hDC.StartDoc(tai_lieu.name or "In PDF")
        try:
            for so_trang in danh_sach_trang:
                trang = tai_lieu[so_trang - 1]
                pix = trang.get_pixmap(matrix=ma_tran)
                anh = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

                # Co giãn ảnh vừa khít vùng in thực tế của máy in, giữ đúng tỉ lệ
                # gốc (không bị méo), căn giữa trang.
                ty_le = min(rong_in / anh.width, cao_in / anh.height)
                w_ve = max(1, int(anh.width * ty_le))
                h_ve = max(1, int(anh.height * ty_le))
                x0 = (rong_in - w_ve) // 2
                y0 = (cao_in - h_ve) // 2

                hDC.StartPage()
                dib = ImageWin.Dib(anh)
                dib.draw(hDC.GetHandleOutput(), (x0, y0, x0 + w_ve, y0 + h_ve))
                hDC.EndPage()
        except Exception as e:
            try:
                hDC.AbortDoc()
            except Exception:
                pass
            return f"{type(e).__name__}: {e}"

        hDC.EndDoc()
        return None

    except Exception as e:
        return f"{type(e).__name__}: {e}"
    finally:
        if hDC is not None:
            try:
                hDC.DeleteDC()
            except Exception:
                pass
        if hprinter is not None:
            try:
                win32print.ClosePrinter(hprinter)
            except Exception:
                pass


def parse_danh_sach_so(chuoi: str, ten_de_bao_loi: str = "danh sách") -> list[int]:
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
    return [int(phan) if phan.isdigit() else phan.lower() for phan in re.split(r"(\d+)", chuoi)]


# =============================================================================
# WIDGET HỖ TRỢ KÉO THẢ DANH SÁCH (THAY THẾ EVENT KÉO THẢ THỦ CÔNG)
# =============================================================================
class SortableListWidget(QListWidget):
    """
    QListWidget hỗ trợ kéo thả tiện lợi. Ghi đè phương thức dropEvent để 
    tự động cập nhật lại số thứ tự (1, 2, 3...) sau khi người dùng kéo thả.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setStyleSheet("font-size: 13px; padding: 5px;")

    def dropEvent(self, event):
        super().dropEvent(event)
        self.update_numbers()

    def update_numbers(self):
        for i in range(self.count()):
            item = self.item(i)
            text = item.text()
            clean_text = re.sub(r'^\d+\.\s*', '', text)
            item.setText(f"{i + 1}. {clean_text}")


# =============================================================================
# ỨNG DỤNG PYQT6
# =============================================================================
class InPageApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.thu_muc_goc = None
        self.danh_sach_file = []
        self.stop_event = threading.Event()
        self.log_queue = queue.Queue()

        self.loai_file = None
        self.che_do_sheet = "vi_tri"
        self.vi_tri_sheet = 1
        self.danh_sach_vi_tri_sheet = "1"
        self.ten_sheet = ""
        self.danh_sach_trang_excel = "1"
        self.che_do_trang_pdf = "toan_bo"
        self.danh_sach_trang_pdf = "1"
        self.so_ban_in = 1

        self._dang_in = False
        self._worker_thread = None
        self._dang_cho_dong_cua_so = False
        self.danh_sach_loi = []
        self.poll_timer = QTimer(self)
        self.poll_timer.timeout.connect(self._poll_queue)

        self._check_thread_timer = QTimer(self)
        self._check_thread_timer.timeout.connect(self._doi_worker_roi_dong)

        self.setWindowTitle(TEN_PROJECT)
        self.resize(700, 650)
        self.setMinimumSize(600, 500)

        self._don_dep_file_tam_cu()
        self._build_folder_select_ui()

    def _don_dep_file_tam_cu(self):
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
                    pass
        except Exception:
            pass

    # ---------- Các phương thức hỗ trợ Layout ----------
    def tao_hang_ngang(self, *widgets):
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        for w in widgets:
            if isinstance(w, int):
                layout.addSpacing(w)
            elif w == "stretch":
                layout.addStretch()
            else:
                layout.addWidget(w)
        container = QWidget()
        container.setLayout(layout)
        return container

    # ---------- Giao diện chọn thư mục ----------
    def _build_folder_select_ui(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        title = QLabel(TEN_PROJECT)
        title.setFont(QFont("Segoe UI", 15, QFont.Weight.Bold))
        title.setStyleSheet("color: #1565c0;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        subtitle = QLabel("Chọn thư mục cần quét")
        subtitle.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(subtitle)

        desc = QLabel("Chương trình sẽ quét thư mục bạn chọn (bao gồm cả các thư mục con bên trong).\nBước tiếp theo bạn sẽ chọn quét file Excel hay file PDF.")
        desc.setStyleSheet("color: #666;")
        desc.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc.setWordWrap(True)
        layout.addWidget(desc)

        self.path_label = QLabel("(Chưa chọn thư mục)")
        self.path_label.setStyleSheet("color: #2e7d32; font-size: 13px;")
        self.path_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.path_label.setWordWrap(True)
        layout.addWidget(self.path_label)

        btn_select = QPushButton("Chọn thư mục...")
        btn_select.setFixedSize(200, 40)
        btn_select.setStyleSheet("background-color: #1565c0; color: white; font-weight: bold; font-size: 13px; border-radius: 5px;")
        btn_select.clicked.connect(self.chon_thu_muc)
        layout.addWidget(btn_select, alignment=Qt.AlignmentFlag.AlignHCenter)

        self.btn_tiep_tuc = QPushButton("Tiếp tục")
        self.btn_tiep_tuc.setFixedSize(200, 40)
        self.btn_tiep_tuc.setStyleSheet("""
            QPushButton { background-color: #2e7d32; color: white; font-weight: bold; font-size: 13px; border-radius: 5px; }
            QPushButton:disabled { background-color: #a5d6a7; color: #f1f1f1; }
        """)
        self.btn_tiep_tuc.setEnabled(self.thu_muc_goc is not None)
        self.btn_tiep_tuc.clicked.connect(self._xac_nhan_thu_muc)
        layout.addWidget(self.btn_tiep_tuc, alignment=Qt.AlignmentFlag.AlignHCenter)

        layout.addStretch()
        self.setCentralWidget(widget)

    def chon_thu_muc(self):
        duong_dan = QFileDialog.getExistingDirectory(self, "Chọn thư mục chứa file Excel/PDF cần in")
        if duong_dan:
            self.thu_muc_goc = Path(duong_dan)
            self.path_label.setText(str(self.thu_muc_goc))
            self.btn_tiep_tuc.setEnabled(True)

    def _xac_nhan_thu_muc(self):
        if not self.thu_muc_goc:
            QMessageBox.warning(self, "Chưa chọn thư mục", "Vui lòng chọn thư mục trước khi tiếp tục.")
            return
        self._build_loai_file_ui()

    # ---------- Giao diện chọn loại file ----------
    def _build_loai_file_ui(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        title = QLabel("Chọn loại file cần in")
        title.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        path_lbl = QLabel(f"Thư mục: {self.thu_muc_goc}")
        path_lbl.setStyleSheet("color: #666;")
        path_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        path_lbl.setWordWrap(True)
        layout.addWidget(path_lbl)

        desc = QLabel("Chỉ chọn được 1 loại — sau khi chọn, loại còn lại sẽ bị bỏ qua\nhoàn toàn (không quét, không hiện cài đặt liên quan).")
        desc.setStyleSheet("color: #666;")
        desc.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc.setWordWrap(True)
        layout.addWidget(desc)

        layout.addSpacing(10)
        
        btn_excel = QPushButton("📊  File Excel\n(.xls / .xlsx)")
        btn_excel.setFixedSize(160, 80)
        btn_excel.setStyleSheet("background-color: #2e7d32; color: white; font-weight: bold; font-size: 14px; border-radius: 8px;")
        btn_excel.clicked.connect(lambda: self._chon_loai_file("excel"))
        
        btn_pdf = QPushButton("📄  File PDF\n(.pdf)")
        btn_pdf.setFixedSize(160, 80)
        btn_pdf.setStyleSheet("background-color: #c62828; color: white; font-weight: bold; font-size: 14px; border-radius: 8px;")
        btn_pdf.clicked.connect(lambda: self._chon_loai_file("pdf"))

        btn_row = self.tao_hang_ngang("stretch", btn_excel, 20, btn_pdf, "stretch")
        layout.addWidget(btn_row)
        layout.addSpacing(20)

        btn_back = QPushButton("Đổi thư mục")
        btn_back.setFixedSize(120, 35)
        btn_back.clicked.connect(self._build_folder_select_ui)
        layout.addWidget(btn_back, alignment=Qt.AlignmentFlag.AlignHCenter)

        layout.addStretch()
        self.setCentralWidget(widget)

    def _chon_loai_file(self, loai):
        self.loai_file = loai
        self._quet_thu_muc()
        self._build_settings_ui()

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
        danh_sach_file.sort(key=lambda p: khoa_sap_xep_tu_nhien(str(p.relative_to(self.thu_muc_goc))))
        self.danh_sach_file = danh_sach_file

    # ---------- Giao diện Cài đặt ----------
    def _build_settings_ui(self):
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("QScrollArea { border: none; }")
        
        content_widget = QWidget()
        scroll_area.setWidget(content_widget)
        layout = QVBoxLayout(content_widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        ten_loai = "Excel" if self.loai_file == "excel" else "PDF"
        
        title = QLabel(f"Cài đặt in — File {ten_loai}")
        title.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        layout.addWidget(title)

        path_lbl = QLabel(f"Thư mục: {self.thu_muc_goc}")
        path_lbl.setStyleSheet("color: #666;")
        path_lbl.setWordWrap(True)
        layout.addWidget(path_lbl)

        self.loi_label = QLabel("")
        self.loi_label.setStyleSheet("color: #c62828;")
        self.loi_label.setWordWrap(True)

        if self.loai_file == "excel":
            lbl_sheet = QLabel("Sheet cần in:")
            lbl_sheet.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            layout.addWidget(lbl_sheet)

            self.bg_sheet = QButtonGroup(self)
            
            self.rb_vitri = QRadioButton("Theo số thứ tự sheet:")
            self.vi_tri_var = QLineEdit(str(self.vi_tri_sheet))
            self.vi_tri_var.setFixedWidth(50)
            lbl_vitri_hint = QLabel("(1 = sheet đầu tiên)")
            lbl_vitri_hint.setStyleSheet("color: #666;")
            layout.addWidget(self.tao_hang_ngang(self.rb_vitri, self.vi_tri_var, 5, lbl_vitri_hint, "stretch"))

            self.rb_danhsach = QRadioButton("Theo danh sách/khoảng sheet:")
            self.danh_sach_var = QLineEdit(self.danh_sach_vi_tri_sheet)
            self.danh_sach_var.setFixedWidth(120)
            lbl_danhsach_hint = QLabel("(vd: 1,2  hoặc  3-4  hoặc  1-2,4)")
            lbl_danhsach_hint.setStyleSheet("color: #666;")
            layout.addWidget(self.tao_hang_ngang(self.rb_danhsach, self.danh_sach_var, 5, lbl_danhsach_hint, "stretch"))

            self.rb_ten = QRadioButton("Theo tên sheet:")
            self.ten_var = QLineEdit(self.ten_sheet)
            self.ten_var.setFixedWidth(150)
            lbl_ten_hint = QLabel("(áp dụng cho mọi file)")
            lbl_ten_hint.setStyleSheet("color: #666;")
            layout.addWidget(self.tao_hang_ngang(self.rb_ten, self.ten_var, 5, lbl_ten_hint, "stretch"))

            self.rb_tatca = QRadioButton("In tất cả các sheet trong file")
            layout.addWidget(self.rb_tatca)

            self.bg_sheet.addButton(self.rb_vitri, 1)
            self.bg_sheet.addButton(self.rb_danhsach, 2)
            self.bg_sheet.addButton(self.rb_ten, 3)
            self.bg_sheet.addButton(self.rb_tatca, 4)

            if self.che_do_sheet == "vi_tri": self.rb_vitri.setChecked(True)
            elif self.che_do_sheet == "danh_sach": self.rb_danhsach.setChecked(True)
            elif self.che_do_sheet == "ten": self.rb_ten.setChecked(True)
            elif self.che_do_sheet == "tat_ca": self.rb_tatca.setChecked(True)

            layout.addSpacing(10)
            lbl_trang = QLabel("Trang cần in trong mỗi sheet:")
            lbl_trang.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            layout.addWidget(lbl_trang)

            self.trang_excel_var = QLineEdit(self.danh_sach_trang_excel)
            self.trang_excel_var.setFixedWidth(150)
            lbl_trang_hint = QLabel("(vd: 1,2  hoặc  3-4  hoặc  1,3,5-7)")
            lbl_trang_hint.setStyleSheet("color: #666;")
            layout.addWidget(self.tao_hang_ngang(self.trang_excel_var, 5, lbl_trang_hint, "stretch"))

            lbl_warn = QLabel("⚠ Nhớ kiểm tra lại đúng nội dung trong các ô nhập ở trên (sheet lẫn\ntrang) trước khi bấm Tiếp tục — chương trình luôn in đúng theo nội\ndung đang hiển thị trong ô, kể cả khi bạn quên sửa lại.")
            lbl_warn.setStyleSheet("color: #c62828;")
            layout.addWidget(lbl_warn)

        else: # PDF
            lbl_pdf = QLabel("Trang cần in:")
            lbl_pdf.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            layout.addWidget(lbl_pdf)

            self.bg_pdf = QButtonGroup(self)
            self.rb_pdf_toanbo = QRadioButton("In toàn bộ trang trong file")
            self.rb_pdf_chidinh = QRadioButton("Chỉ in trang chỉ định:")
            self.bg_pdf.addButton(self.rb_pdf_toanbo, 1)
            self.bg_pdf.addButton(self.rb_pdf_chidinh, 2)

            if self.che_do_trang_pdf == "toan_bo": self.rb_pdf_toanbo.setChecked(True)
            else: self.rb_pdf_chidinh.setChecked(True)

            layout.addWidget(self.rb_pdf_toanbo)
            
            self.trang_pdf_var = QLineEdit(self.danh_sach_trang_pdf)
            self.trang_pdf_var.setFixedWidth(150)
            layout.addWidget(self.tao_hang_ngang(self.rb_pdf_chidinh, 10, self.trang_pdf_var, "stretch"))

            lbl_pdf_hint = QLabel("(vd: 1,2  hoặc  3-4  hoặc  1,3,5-7)")
            lbl_pdf_hint.setStyleSheet("color: #666;")
            layout.addWidget(self.tao_hang_ngang(30, lbl_pdf_hint, "stretch"))

            lbl_warn_pdf = QLabel("⚠ Nhớ kiểm tra lại đúng số trang trong ô trên trước khi bấm Tiếp tục —\nnếu chọn \"Chỉ in trang chỉ định\" mà quên sửa ô nhập, chương trình sẽ\nin đúng theo nội dung đang hiện trong ô đó.")
            lbl_warn_pdf.setStyleSheet("color: #c62828;")
            layout.addWidget(lbl_warn_pdf)

        layout.addSpacing(10)
        lbl_ban = QLabel("Số bản in:")
        lbl_ban.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        layout.addWidget(lbl_ban)

        self.so_ban_var = QLineEdit(str(self.so_ban_in))
        self.so_ban_var.setFixedWidth(60)
        lbl_ban_hint = QLabel("bản mỗi file")
        lbl_ban_hint.setStyleSheet("color: #666;")
        layout.addWidget(self.tao_hang_ngang(self.so_ban_var, 8, lbl_ban_hint, "stretch"))

        layout.addWidget(self.loi_label)

        # Hàng nút bấm
        btn_back = QPushButton("Đổi loại file")
        btn_back.setFixedSize(120, 35)
        btn_back.clicked.connect(self._build_loai_file_ui)

        btn_next = QPushButton("Tiếp tục")
        btn_next.setFixedSize(120, 35)
        btn_next.setStyleSheet("background-color: #2e7d32; color: white; font-weight: bold;")
        btn_next.clicked.connect(self._xac_nhan_cai_dat)

        layout.addWidget(self.tao_hang_ngang(btn_back, 10, btn_next, "stretch"))
        layout.addStretch()

        self.setCentralWidget(scroll_area)

    def _xac_nhan_cai_dat(self):
        self.loi_label.setText("")

        if self.loai_file == "excel":
            if self.rb_vitri.isChecked():
                che_do = "vi_tri"
                try:
                    vi_tri = int(self.vi_tri_var.text())
                    if vi_tri < 1: raise ValueError
                except ValueError:
                    self.loi_label.setText("Số thứ tự sheet phải là số nguyên dương (VD: 1, 2, 3...).")
                    return
                self.vi_tri_sheet = vi_tri

            elif self.rb_danhsach.isChecked():
                che_do = "danh_sach"
                try:
                    parse_danh_sach_so(self.danh_sach_var.text())
                except ValueError as e:
                    self.loi_label.setText(f"Danh sách sheet không hợp lệ: {e}")
                    return
                self.danh_sach_vi_tri_sheet = self.danh_sach_var.text().strip()

            elif self.rb_ten.isChecked():
                che_do = "ten"
                ten = self.ten_var.text().strip()
                if not ten:
                    self.loi_label.setText("Vui lòng nhập tên sheet.")
                    return
                self.ten_sheet = ten
            else:
                che_do = "tat_ca"

            try:
                parse_danh_sach_so(self.trang_excel_var.text())
            except ValueError as e:
                self.loi_label.setText(f"Trang cần in không hợp lệ: {e}")
                return

            self.che_do_sheet = che_do
            self.danh_sach_trang_excel = self.trang_excel_var.text().strip()

        else:
            if self.rb_pdf_toanbo.isChecked():
                che_do_pdf = "toan_bo"
            else:
                che_do_pdf = "chi_dinh"
                try:
                    parse_danh_sach_so(self.trang_pdf_var.text())
                except ValueError as e:
                    self.loi_label.setText(f"Trang cần in không hợp lệ: {e}")
                    return
            self.che_do_trang_pdf = che_do_pdf
            self.danh_sach_trang_pdf = self.trang_pdf_var.text().strip()

        try:
            so_ban = int(self.so_ban_var.text())
            if so_ban < 1: raise ValueError
        except ValueError:
            self.loi_label.setText("Số bản in phải là số nguyên dương (VD: 1, 2, 3...).")
            return
        
        self.so_ban_in = so_ban
        self._build_confirm_ui()

    # ---------- Giao diện Xác nhận ----------
    def _build_confirm_ui(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(20, 20, 20, 20)

        so_file = len(self.danh_sach_file)

        if self.loai_file == "excel":
            if self.che_do_sheet == "vi_tri": mo_ta_sheet = f"sheet số {self.vi_tri_sheet}"
            elif self.che_do_sheet == "danh_sach": mo_ta_sheet = f"các sheet số {self.danh_sach_vi_tri_sheet}"
            elif self.che_do_sheet == "ten": mo_ta_sheet = f"sheet tên \"{self.ten_sheet}\""
            else: mo_ta_sheet = "tất cả các sheet"

            mo_ta_tom_tat = f"Sẽ gửi lệnh in {so_file} file Excel\nIn {mo_ta_sheet}, trang {self.danh_sach_trang_excel} mỗi sheet, {self.so_ban_in} bản mỗi file"
        else:
            mo_ta_pdf = "in toàn bộ trang" if self.che_do_trang_pdf == "toan_bo" else f"chỉ in trang {self.danh_sach_trang_pdf}"
            mo_ta_tom_tat = f"Sẽ gửi lệnh in {so_file} file PDF\n{mo_ta_pdf}, {self.so_ban_in} bản mỗi file"

        lbl_summary = QLabel(mo_ta_tom_tat)
        lbl_summary.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        lbl_summary.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_summary)
        layout.addSpacing(10)

        if so_file == 0:
            ten_loai = "Excel" if self.loai_file == "excel" else "PDF"
            lbl_err = QLabel(f"Không tìm thấy file {ten_loai} nào trong thư mục này.")
            lbl_err.setStyleSheet("color: #c62828;")
            lbl_err.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(lbl_err)

            btn_back = QPushButton("Quay lại")
            btn_back.setFixedSize(120, 35)
            btn_back.clicked.connect(self._build_settings_ui)
            layout.addWidget(btn_back, alignment=Qt.AlignmentFlag.AlignHCenter)
            layout.addStretch()
            self.setCentralWidget(widget)
            return

        lbl_hint = QLabel("Mẹo: nhấn giữ chuột và kéo thả 1 dòng để sắp xếp lại thứ tự in.")
        lbl_hint.setStyleSheet("color: #666;")
        layout.addWidget(lbl_hint)

        self.listbox = SortableListWidget()
        for idx, p in enumerate(self.danh_sach_file, start=1):
            try:
                duong_dan_hien_thi = str(p.relative_to(self.thu_muc_goc))
            except ValueError:
                duong_dan_hien_thi = str(p)
            item = QListWidgetItem(f"{idx}. {duong_dan_hien_thi}")
            item.setData(Qt.ItemDataRole.UserRole, p)
            self.listbox.addItem(item)
            
        layout.addWidget(self.listbox)

        # Hàng nút bấm
        btn_start = QPushButton("Bắt đầu in")
        btn_start.setFixedSize(120, 35)
        btn_start.setStyleSheet("background-color: #2e7d32; color: white; font-weight: bold;")
        btn_start.clicked.connect(self._start_printing)

        btn_back = QPushButton("Quay lại")
        btn_back.setFixedSize(120, 35)
        btn_back.clicked.connect(self._build_settings_ui)

        btn_cancel = QPushButton("Hủy bỏ")
        btn_cancel.setFixedSize(120, 35)
        btn_cancel.clicked.connect(self.close)

        layout.addWidget(self.tao_hang_ngang("stretch", btn_start, 10, btn_back, 10, btn_cancel, "stretch"))
        self.setCentralWidget(widget)

    # ---------- Giao diện khi in ----------
    def _build_progress_ui(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(20, 20, 20, 20)

        lbl = QLabel("Đang gửi lệnh in...")
        lbl.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        layout.addWidget(lbl)

        self.progress = QProgressBar()
        self.progress.setMaximum(len(self.danh_sach_file))
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        self.log_box = QListWidget()
        self.log_box.setStyleSheet("font-size: 12px; padding: 5px;")
        layout.addWidget(self.log_box, stretch=2)

        # Khung riêng LIỆT KÊ LỖI — tách hẳn khỏi log chung để không bị bỏ sót
        # giữa hàng trăm dòng "đã in thành công". Chỉ hiện những file/sheet có
        # lỗi, đỏ nổi bật, cập nhật ngay khi lỗi phát sinh (không cần đợi in
        # xong hết mới biết).
        self.nhan_khung_loi = QLabel("⚠ File/sheet bị lỗi (0):")
        self.nhan_khung_loi.setStyleSheet("color: #c62828; font-weight: bold;")
        layout.addWidget(self.nhan_khung_loi)

        self.error_box = QListWidget()
        self.error_box.setStyleSheet(
            "QListWidget { background-color: #fdecea; border: 1px solid #c62828; "
            "font-size: 12px; padding: 5px; color: #b71c1c; }"
        )
        self.error_box.setMaximumHeight(140)
        layout.addWidget(self.error_box, stretch=1)

        self.cancel_btn = QPushButton("HỦY TẤT CẢ LỆNH IN")
        self.cancel_btn.setFixedHeight(40)
        self.cancel_btn.setStyleSheet("""
            QPushButton { background-color: #c62828; color: white; font-weight: bold; }
            QPushButton:disabled { background-color: #e57373; color: #f1f1f1; }
        """)
        self.cancel_btn.clicked.connect(self._cancel_printing)
        layout.addWidget(self.cancel_btn)

        self.setCentralWidget(widget)

    def _start_printing(self):
        # Cập nhật danh sách file dựa trên thứ tự người dùng đã kéo thả trên QListWidget
        self.danh_sach_file = [self.listbox.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.listbox.count())]

        self.stop_event.clear()
        self._dang_in = True
        self.danh_sach_loi = []  # reset danh sách lỗi mỗi lượt in mới
        self._build_progress_ui()
        
        self._worker_thread = threading.Thread(target=self._print_worker, daemon=True)
        self._worker_thread.start()
        
        self.poll_timer.start(100)

    def _cancel_printing(self):
        self.stop_event.set()
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.setText("Đang hủy, chờ file hiện tại in xong...")

    def _poll_queue(self):
        if self._dang_cho_dong_cua_so:
            return 
        try:
            while True:
                msg = self.log_queue.get_nowait()
                if msg[0] == "progress":
                    _, idx, ten_file, trang_thai = msg
                    self.progress.setValue(idx)
                    self.status_label.setText(f"[{idx}/{len(self.danh_sach_file)}] {ten_file}")
                    dong_log = f"{ten_file}  ->  {trang_thai}"
                    self.log_box.addItem(dong_log)
                    if self.log_box.count() > 500:
                        self.log_box.takeItem(0)
                    self.log_box.scrollToBottom()

                    # Nhận diện dòng có lỗi (không phân biệt hoa/thường, vì lỗi cấp
                    # sheet Excel dùng chữ "lỗi" thường, lỗi cấp file dùng "Lỗi" hoa)
                    # và đẩy riêng vào khung báo lỗi để không bị chìm trong log chung.
                    if "lỗi" in trang_thai.lower() or "không có sheet nào" in trang_thai.lower():
                        self.danh_sach_loi.append(dong_log)
                        self.error_box.addItem(dong_log)
                        self.error_box.scrollToBottom()
                        self.nhan_khung_loi.setText(f"⚠ File/sheet bị lỗi ({len(self.danh_sach_loi)}):")

                elif msg[0] == "done":
                    _, thanh_cong, loi, huy = msg
                    self._dang_in = False
                    self.poll_timer.stop()
                    self._show_summary(thanh_cong, loi, huy)
                    return
        except queue.Empty:
            pass

    # ---------- Giao diện Hoàn thành ----------
    def _show_summary(self, thanh_cong, loi, huy):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        tieu_de = "Đã hủy" if huy else "Hoàn thành"
        mau = "#c62828" if huy else "#2e7d32"
        
        lbl_title = QLabel(tieu_de)
        lbl_title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        lbl_title.setStyleSheet(f"color: {mau};")
        lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_title)

        lbl_stats = QLabel(f"Lệnh in thành công: {thanh_cong}\nLệnh in bị lỗi: {loi}")
        lbl_stats.setFont(QFont("Segoe UI", 11))
        lbl_stats.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_stats)

        # Nếu có lỗi, hiện lại chi tiết ngay tại đây (không chỉ mỗi con số) —
        # vì màn tiến trình đã bị thay thế, đây là nơi cuối cùng người dùng có
        # thể xem lại chính xác file/sheet nào lỗi và lý do gì.
        if self.danh_sach_loi:
            lbl_loi_tieu_de = QLabel(f"⚠ Chi tiết {len(self.danh_sach_loi)} lỗi:")
            lbl_loi_tieu_de.setStyleSheet("color: #c62828; font-weight: bold;")
            layout.addWidget(lbl_loi_tieu_de)

            hop_loi = QListWidget()
            hop_loi.addItems(self.danh_sach_loi)
            hop_loi.setStyleSheet(
                "QListWidget { background-color: #fdecea; border: 1px solid #c62828; "
                "font-size: 12px; padding: 5px; color: #b71c1c; }"
            )
            hop_loi.setMaximumHeight(160)
            layout.addWidget(hop_loi)

        if huy:
            lbl_huy_warn = QLabel("Lưu ý: những file đã gửi lệnh in TRƯỚC KHI bấm hủy\nvẫn có thể đang nằm trong hàng đợi máy in (Print Queue)\nvì lệnh in đã được gửi đi rồi. Hãy mở mục quản lý máy in\ntrong Windows để xóa các lệnh in đó nếu cần.")
            lbl_huy_warn.setStyleSheet("color: #c62828;")
            layout.addWidget(lbl_huy_warn)

        co_file_pdf = any(p.suffix.lower() == ".pdf" for p in self.danh_sach_file)
        if co_file_pdf:
            lbl_pdf_warn = QLabel("Lưu ý về file PDF: chương trình in trực tiếp qua GDI của Windows,\nkhông còn thông qua Edge/Acrobat hay bất kỳ trình đọc PDF nào nữa.\nChương trình không thể xác nhận việc in vật lý đã hoàn tất 100%\n(mực đã ra giấy) — vui lòng kiểm tra thực tế bản in.")
            lbl_pdf_warn.setStyleSheet("color: #666;")
            layout.addWidget(lbl_pdf_warn)

        btn_restart = QPushButton("Quét thư mục khác")
        btn_restart.setFixedSize(150, 35)
        btn_restart.clicked.connect(self._build_folder_select_ui)

        btn_close = QPushButton("Đóng")
        btn_close.setFixedSize(120, 35)
        btn_close.clicked.connect(self.close)

        layout.addWidget(self.tao_hang_ngang("stretch", btn_restart, 10, btn_close, "stretch"))
        layout.addStretch()
        self.setCentralWidget(widget)

    # ---------- Các thao tác Window cơ bản & an toàn ----------
    def closeEvent(self, event: QCloseEvent):
        if self._dang_in:
            reply = QMessageBox.question(
                self, "Đang in dở",
                "Chương trình đang gửi lệnh in.\n"
                "Nếu thoát ngay bây giờ, các lệnh in đã gửi vẫn nằm trong hàng đợi máy in,\n"
                "nhưng chương trình sẽ dừng gửi thêm lệnh mới.\n\n"
                "Bạn có chắc muốn thoát không?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.stop_event.set()
                self._dang_cho_dong_cua_so = True
                self.poll_timer.stop()
                self._vo_hieu_hoa_cua_so_dang_dong()
                self._check_thread_timer.start(200)
                event.ignore()
            else:
                event.ignore()
        else:
            event.accept()

    def _vo_hieu_hoa_cua_so_dang_dong(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        lbl = QLabel("Đang dừng và dọn dẹp, vui lòng chờ...")
        lbl.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl)
        self.setCentralWidget(widget)

    def _doi_worker_roi_dong(self):
        if self._worker_thread is not None and self._worker_thread.is_alive():
            return
        self._check_thread_timer.stop()
        QApplication.quit()

    # ---------- Xử lý in thực tế (PDF & Excel) ----------
    def _in_pdf(self, duong_dan_file: Path, so_ban_in: int):
        """
        In trực tiếp file PDF qua GDI của Windows (dùng PyMuPDF render từng
        trang thành ảnh rồi gửi thẳng cho driver máy in) — không còn thông qua
        bất kỳ ứng dụng đọc PDF nào, không còn cần tạo/xóa file tạm.
        """
        try:
            tai_lieu = fitz.open(str(duong_dan_file))
        except Exception as e:
            return (f"Lỗi: không mở được file PDF ({e})", 0, 1)

        try:
            tong_trang = tai_lieu.page_count

            if self.che_do_trang_pdf == "toan_bo":
                danh_sach_trang = list(range(1, tong_trang + 1))
                trang_vuot_qua = []
            else:
                danh_sach_trang_nhap = parse_danh_sach_so(self.danh_sach_trang_pdf, "Danh sách trang PDF")
                danh_sach_trang = [t for t in danh_sach_trang_nhap if t <= tong_trang]
                trang_vuot_qua = [t for t in danh_sach_trang_nhap if t > tong_trang]
                if not danh_sach_trang:
                    return (
                        f"Lỗi: file chỉ có {tong_trang} trang, không có trang nào trong "
                        f"{danh_sach_trang_nhap} hợp lệ", 0, 1
                    )

            # Phòng trường hợp file PDF rỗng/hỏng (0 trang) — không có bước kiểm
            # tra này thì tai_lieu[danh_sach_trang[0] - 1] bên dưới sẽ ném
            # IndexError khó hiểu thay vì 1 dòng lỗi rõ ràng.
            if not danh_sach_trang:
                return ("Lỗi: file PDF không có trang nào (có thể bị rỗng hoặc hỏng)", 0, 1)

            la_ngang = kiem_tra_pdf_huong_ngang(tai_lieu[danh_sach_trang[0] - 1])

            loi = in_truc_tiep_qua_gdi(tai_lieu, danh_sach_trang, so_ban_in, la_ngang)
            if loi:
                return (f"Lỗi khi in: {loi}", 0, 1)

            if self.che_do_trang_pdf == "toan_bo":
                mo_ta = f"Đã in toàn bộ {tong_trang} trang (x{so_ban_in} bản)"
            else:
                mo_ta = f"Đã in trang {', '.join(str(t) for t in danh_sach_trang)} (x{so_ban_in} bản)"
                if trang_vuot_qua:
                    mo_ta += f" (bỏ qua trang {', '.join(str(t) for t in trang_vuot_qua)} vì file chỉ có {tong_trang} trang)"
            return (mo_ta, 1, 0)

        except Exception as e:
            return (f"Lỗi khi in PDF: {e}", 0, 1)
        finally:
            tai_lieu.close()

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
                loi_mo_excel = str(e)

            try:
                khoang_trang_excel = gom_nhom_lien_tuc(parse_danh_sach_so(self.danh_sach_trang_excel, "Trang Excel"))
            except ValueError as e:
                loi_trang_excel = str(e)

        for idx, duong_dan_file in enumerate(self.danh_sach_file, start=1):
            if self.stop_event.is_set():
                bi_huy = True
                break

            if self.loai_file == "pdf":
                mo_ta, thanh_cong, loi = self._in_pdf(duong_dan_file, self.so_ban_in)
                so_sheet_thanh_cong += thanh_cong
                so_sheet_loi += loi
                self.log_queue.put(("progress", idx, duong_dan_file.name, mo_ta))
                continue

            if excel is None:
                so_sheet_loi += 1
                self.log_queue.put(("progress", idx, duong_dan_file.name, f"Lỗi: không mở được Excel ({loi_mo_excel})"))
                continue

            if loi_trang_excel is not None:
                so_sheet_loi += 1
                self.log_queue.put(("progress", idx, duong_dan_file.name, f"Lỗi: trang Excel không hợp lệ ({loi_trang_excel})"))
                continue

            wb = None
            ket_qua_sheets = []
            try:
                wb = excel.Workbooks.Open(str(duong_dan_file), ReadOnly=True)

                if self.che_do_sheet == "vi_tri": danh_sach_ws = [("vi_tri", self.vi_tri_sheet)]
                elif self.che_do_sheet == "danh_sach": danh_sach_ws = [("vi_tri", vt) for vt in parse_danh_sach_so(self.danh_sach_vi_tri_sheet)]
                elif self.che_do_sheet == "ten": danh_sach_ws = [("ten", self.ten_sheet)]
                else: danh_sach_ws = [("obj", ws) for ws in wb.Sheets]

                for loai, tham_chieu in danh_sach_ws:
                    ws = None
                    try:
                        ws = tham_chieu if loai == "obj" else wb.Sheets(tham_chieu)
                        for bat_dau, ket_thuc in khoang_trang_excel:
                            ws.PrintOut(From=bat_dau, To=ket_thuc, Copies=self.so_ban_in)

                        so_sheet_thanh_cong += 1
                        trang_da_in = ",".join(f"{a}" if a == b else f"{a}-{b}" for a, b in khoang_trang_excel)
                        ket_qua_sheets.append(f"{ws.Name}: đã in trang {trang_da_in} (x{self.so_ban_in} bản)")
                    except Exception as e:
                        so_sheet_loi += 1
                        ten_hien_thi = tham_chieu if loai != "obj" else "?"
                        ket_qua_sheets.append(f"Sheet {ten_hien_thi}: lỗi - {e}")
                    finally:
                        ws = None

            except Exception as e:
                so_sheet_loi += 1
                ket_qua_sheets.append(f"Không mở được file: {e}")
            finally:
                danh_sach_ws = None
                if wb is not None:
                    try: wb.Close(SaveChanges=False)
                    except Exception: pass
                wb = None

            trang_thai = "; ".join(ket_qua_sheets) if ket_qua_sheets else "Không có sheet nào được in"
            self.log_queue.put(("progress", idx, duong_dan_file.name, trang_thai))
            if idx % 5 == 0:
                gc.collect()

        if excel is not None:
            try: excel.Quit()
            except Exception: pass
            excel = None
            gc.collect()
            
        pythoncom.CoUninitialize()

        self.log_queue.put(("done", so_sheet_thanh_cong, so_sheet_loi, bi_huy))


if __name__ == "__main__":
    from PyQt6.QtGui import QPalette, QColor
    
    app = QApplication(sys.argv)
    app.setStyle("Fusion") 
    
    # Ép dùng Light Theme bằng QPalette
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(250, 250, 250))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(0, 0, 0))
    palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(245, 245, 245))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(0, 0, 0))
    palette.setColor(QPalette.ColorRole.Text, QColor(0, 0, 0))
    palette.setColor(QPalette.ColorRole.Button, QColor(240, 240, 240))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(0, 0, 0))
    palette.setColor(QPalette.ColorRole.BrightText, QColor(255, 0, 0))
    palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
    app.setPalette(palette)
    
    # Ép CSS bổ sung cho chắc cốp mấy cái hộp thoại và background
    app.setStyleSheet("""
        QMainWindow, QMessageBox, QDialog { background-color: #ffffff; color: #000000; }
        QLabel, QRadioButton { color: #333333; }
        QListWidget { background-color: #ffffff; color: #000000; border: 1px solid #ccc; }
        QLineEdit { background-color: #ffffff; color: #000000; border: 1px solid #ccc; padding: 2px; }
    """)

    window = InPageApp()
    window.show()
    sys.exit(app.exec())

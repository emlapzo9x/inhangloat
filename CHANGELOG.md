# Changelog — In Hàng Loạt by An Duc

## v1.0 — Bản gốc (console script)
- Script `inpage1-2.py` chạy dòng lệnh thuần, không giao diện.
- Tự động quét thư mục chứa chính file `.py` (và các thư mục con).
- Mở từng file `.xls`/`.xlsx` bằng Excel COM, in cố định trang 1–2 của sheet đầu
  tiên.
- Sửa lỗi tiềm ẩn: crash nếu mở file thất bại (biến `wb` không tồn tại trong
  `finally`), không bỏ qua file khóa tạm `~$...`.

## v1.1 — Chuyển sang giao diện đồ họa (GUI)
- Chuyển từ console sang giao diện `tkinter`.
- Thêm màn hình xác nhận trước khi in: hiển thị danh sách file + tổng số file
  sẽ in.
- Thêm nút **"HỦY TẤT CẢ LỆNH IN"** giữa chừng để tránh in nhầm hàng loạt.
- Chạy luồng in trong thread riêng để không đứng giao diện.

## v1.2 — Chọn sheet linh hoạt
- Cho chọn sheet cần in theo 4 cách: theo số thứ tự, theo danh sách/khoảng số
  thứ tự (vd `1,2` hoặc `3-4`), theo tên sheet, hoặc in tất cả các sheet.
- Sửa trùng lặp code khi xử lý 2 chế độ "vị trí" và "tên" giống hệt nhau.

## v1.3 — Chọn thư mục qua giao diện
- Không cần copy file `.exe` vào chung thư mục Excel nữa — thêm hộp thoại
  duyệt thư mục (`filedialog.askdirectory`) để chọn nơi cần quét.
- Thêm nút "Đổi thư mục" và "Quét thư mục khác" để dùng lại app cho thư mục
  khác mà không cần khởi động lại.
- Sửa lỗi `AttributeError: 'NoneType' object has no attribute 'rglob'` (sót
  đoạn quét thư mục cũ chạy trước khi người dùng kịp chọn thư mục).

## v1.4 — Vá lỗi & tối ưu vòng đời chương trình
- Chặn việc đóng cửa sổ đột ngột (nút X) trong lúc đang in — thêm xác nhận
  trước khi thoát để tránh bỏ sót tiến trình Excel.
- Reset cờ hủy (`stop_event`) mỗi lượt in mới, tránh lượt in sau bị coi là đã
  hủy do cờ cũ còn sót từ lượt trước.
- Gộp code trùng lặp, đổi tên biến cho đúng bản chất (đếm theo sheet).

## v1.5 — Hỗ trợ file PDF, đổi tên dự án
- Đổi tên chương trình thành **"In Hàng Loạt by An Duc"**.
- Hỗ trợ in file `.pdf` song song với Excel: dùng `pypdf` cắt đúng trang cần
  in ra file tạm, gửi cho trình đọc PDF mặc định của Windows để in.

## v1.6 — Hoàn thiện hiển thị danh sách file
- Thêm số thứ tự cho từng file trong danh sách xác nhận.
- Thêm thanh cuộn ngang để xem trọn đường dẫn dài.
- Giải phóng rõ ràng các đối tượng COM (workbook, worksheet) + `gc.collect()`
  để giảm nguy cơ Excel.exe treo ngầm.
- Tự động dọn file PDF tạm cũ (>1 giờ) mỗi lần khởi động, đề phòng lần chạy
  trước bị tắt đột ngột.

## v1.7 — Trang Excel/PDF độc lập, hỗ trợ trang lẻ
- Tách khoảng trang Excel và PDF thành 2 cấu hình độc lập (trước đó dùng
  chung).
- Hỗ trợ chọn trang **không liên tục** (vd `1,3,5,7`) cho cả Excel lẫn PDF,
  không chỉ giới hạn ở một khoảng liên tục.
- Excel: tự động gộp các trang liền kề thành khoảng để giảm số lệnh in gửi
  cho Excel (vì Excel chỉ nhận `From`/`To` liên tục mỗi lệnh).
- Thêm lựa chọn "In toàn bộ trang" cho PDF (gửi thẳng file gốc, không cần cắt
  trang qua file tạm).

## v1.8 — Vá lỗi bảo mật đường dẫn & bố cục giao diện
- **Sửa lỗi MAX_PATH**: bỏ chuỗi số trang khỏi tên file PDF tạm (dễ vượt quá
  260 ký tự nếu chọn nhiều trang lẻ + tên file gốc dài), cắt bớt tên gốc còn
  tối đa 60 ký tự.
- **Sửa lỗi giao diện**: thanh cuộn ngang từng bị đẩy lạc xuống dưới cả hàng
  nút bấm do gắn nhầm vào khung cha — chuyển vào đúng khung chứa danh sách.
- Cho phép resize cửa sổ + tăng kích thước mặc định, tránh nội dung bị bóp
  méo khi màn cài đặt có nhiều dòng hơn.
- Sắp xếp danh sách file theo thứ tự ổn định trước khi in.
- Giảm tần suất gọi `gc.collect()` (mỗi 5 file thay vì mỗi file) để đỡ tốn
  CPU với batch lớn.
- Thêm retry khi xóa file PDF tạm thất bại (thử lại thêm 2 lần).
- Giới hạn tối đa 500 dòng hiển thị trong log để tránh giao diện chậm khi
  quét hàng nghìn file.
- Giải phóng tường minh từng đối tượng sheet (`ws = None`) ngay sau khi in
  xong, thay vì chỉ trông chờ vòng lặp gán lại.

## v1.9 — Sửa lỗi tắt ứng dụng giữa chừng khi đang in
- **Sửa lỗi quan trọng**: khi bấm nút X để thoát lúc đang in và xác nhận vẫn
  thoát, chương trình trước đó đóng cửa sổ ngay lập tức — vì luồng in chạy
  dạng *daemon thread*, tiến trình chính thoát trước khi luồng in kịp chạy
  xong `excel.Quit()` sẽ khiến Excel bị ngắt đột ngột, dễ để sót tiến trình
  Excel chạy ngầm. Giờ chương trình hiển thị màn hình "Đang dừng và dọn dẹp,
  vui lòng chờ..." và đợi luồng in thực sự kết thúc hẳn rồi mới đóng cửa sổ.
- Thêm cờ chặn cập nhật giao diện khi đang trong quá trình chờ đóng, tránh
  lỗi `TclError` do cập nhật nhầm vào widget đã bị hủy.

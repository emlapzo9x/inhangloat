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

## v1.9.1 — Tự sắp xếp thứ tự file + kéo thả tùy chỉnh
- **Sắp xếp "tự nhiên"** danh sách file trước khi in: đúng thứ tự số
  (`1, 2, ..., 10`) thay vì so sánh chuỗi thuần (`1, 10, 2`).
- Thêm khả năng **kéo thả trực tiếp** một dòng trong danh sách ở màn xác nhận
  để tự sắp xếp lại thứ tự in theo ý muốn, không cần đổi tên file.
- Sửa lỗi bố cục: dòng gợi ý "Mẹo: kéo thả..." từng bị hiển thị lạc xuống
  dưới danh sách file do thứ tự `pack()` sai (đúng dạng lỗi từng gặp với
  thanh cuộn ngang ở bản trước) — đã sửa để hiện phía trên danh sách như dự
  định ban đầu.

## v2.0 — Số bản in tùy chỉnh + tối ưu in PDF theo từng trình đọc
- Thêm **số bản in mặc định riêng** cho Excel và PDF (cấu hình ở màn cài đặt).
- Thêm khả năng **chỉnh số bản in riêng cho từng file** ngay trong danh sách
  ở màn xác nhận (chọn 1 file, nhập số, bấm "Áp dụng"), ghi đè số mặc định.
- Excel: dùng thẳng tham số `Copies` của Excel COM để in đúng số bản trong 1
  lệnh.
- Thêm lựa chọn ứng dụng dùng để in PDF: Adobe Acrobat/Reader, Foxit
  Reader/PhantomPDF, hoặc Nitro Pro — các app này hỗ trợ tham số dòng lệnh
  `/t` để in im lặng thật sự, đáng tin cậy hơn verb "print" mặc định của
  Windows.

## v2.1 — Sửa lỗi in PDF ra nhiều bản hơn mong muốn, thêm SumatraPDF
- **Sửa lỗi quan trọng**: PDF bị in ra nhiều bản hơn số bản đã chọn (kể cả khi
  chỉ chọn in 1 bản). Nguyên nhân gốc: các cách in PDF im lặng (verb "print"
  của Windows hay switch dòng lệnh của Acrobat/Foxit/Nitro) không có tham số
  riêng để chỉ định số bản in — chúng luôn in theo đúng số bản đang được lưu
  sẵn trong driver máy in mặc định (giống ô "Copies" khi in tay). Cách cũ (gọi
  lại lệnh in N lần liên tiếp) bị nhân đôi khó lường nếu driver đang "dính"
  sẵn số bản khác 1 từ trước. Sửa bằng cách đặt thẳng số bản vào driver máy in
  (qua `win32print`) ngay trước khi gửi, rồi chỉ gửi **đúng 1 lệnh in duy
  nhất** cho mỗi file — giống cách Excel đã làm đúng từ trước (tham số Copies
  riêng của Excel COM, không phụ thuộc driver nên không dính lỗi này).
- Đặt lại số bản in của driver về 1 sau khi in xong toàn bộ, để không ảnh
  hưởng tới các lần in tay khác của người dùng sau này.
- Thêm **SumatraPDF** vào danh sách ứng dụng hỗ trợ in im lặng qua dòng lệnh
  (cú pháp riêng khác Acrobat/Foxit/Nitro); có dò thêm các đường dẫn cài đặt
  phổ biến vì SumatraPDF hay được cài kiểu portable, không phải lúc nào cũng
  đăng ký trong registry Windows.
- Chuyển màn hình Cài đặt sang dạng **có thể cuộn dọc**, tránh nội dung bị
  bóp méo/cắt cụt khi càng thêm nhiều tùy chọn qua các bản sau này.



## v2.2 — Tách hẳn menu Excel và PDF, bỏ số bản in riêng theo file
- **Tách hoàn toàn luồng Excel và PDF**: thêm màn hình chọn "File Excel" hoặc
  "File PDF" ngay sau khi chọn thư mục — chỉ được chọn 1 loại. Sau khi chọn:
  - Màn cài đặt chỉ hiện đúng các tùy chọn liên quan tới loại đã chọn (chọn
    Excel thì không còn thấy gì về PDF, và ngược lại).
  - Bước quét thư mục cũng chỉ lấy đúng loại file đã chọn — loại còn lại bị
    bỏ qua hoàn toàn ngay từ đầu, không chỉ ẩn trên giao diện.
- **Bỏ tính năng chỉnh số bản in riêng theo từng file** ở màn xác nhận (panel
  "Số bản in cho file đã chọn / Áp dụng" đã bị xóa). Giờ chỉ còn 1 ô "Số bản
  in" duy nhất ở màn cài đặt, áp dụng chung cho mọi file thuộc loại đã chọn.
- Giữ nguyên tính năng kéo thả sắp xếp thứ tự file ở màn xác nhận.

## v2.3 — Bỏ tính năng chọn app in PDF cụ thể, đơn giản hóa cho ổn định
- **Bỏ hoàn toàn tính năng chọn ứng dụng in PDF cụ thể** (Acrobat/Foxit/Nitro/
  SumatraPDF) theo yêu cầu — đây là phần rủi ro nhất trong toàn bộ tính năng
  in PDF (phụ thuộc dò đường dẫn qua registry có thể sai, và cú pháp dòng lệnh
  riêng của từng app, đặc biệt Nitro Pro, chưa được xác nhận chắc chắn đúng
  với mọi phiên bản). Giờ chỉ dùng đúng 1 cách duy nhất: verb "print" mặc định
  của Windows (giống hệt chuột phải file rồi bấm Print).
- **Vẫn giữ nguyên** cơ chế đặt số bản in vào driver máy in qua `win32print`
  (phần đã sửa lỗi "in ra nhiều bản" ở v2.1) — phần này đáng tin cậy, không
  phụ thuộc app bên thứ 3 nào.
- **Phát hiện và sửa luôn 1 lỗi nghiêm trọng đang có sẵn từ v2.2**: do sót khi
  gộp code ở bản trước, hàm `UNG_DUNG_PDF_HO_TRO_IN_LENH` và
  `tim_duong_dan_ung_dung` bị mất định nghĩa nhưng code vẫn gọi tới — gây
  crash `NameError` ngay khi in bất kỳ file PDF nào. Việc bỏ tính năng chọn
  app hôm nay đã giải quyết luôn lỗi này.

## v2.4 — Tự động phát hiện và sửa hướng giấy khi in PDF
- **Sửa lỗi**: file PDF có nội dung nằm ngang bị in ra co nhỏ, lệch sang 1 góc
  tờ giấy dọc (do driver máy in mặc định đang để hướng dọc, và verb "print"
  không tự xoay hướng giấy theo nội dung thật của trang PDF) — trông như file
  "không in được" dù thực chất vẫn có in, chỉ là sai hướng.
- Chương trình giờ tự động đọc kích thước trang PDF thực tế (kể cả trường hợp
  trang bị xoay 90°/270°) để xác định trang đó là dọc hay ngang, rồi đặt đúng
  hướng giấy vào driver máy in TRƯỚC khi gửi từng lệnh in — áp dụng đúng cho
  từng file, kể cả khi thư mục có lẫn cả file dọc và file ngang.
- Giới hạn cần biết: nếu 1 file PDF (không phải khác file, mà ngay bên trong
  1 file) có trang dọc xen trang ngang, chương trình chỉ xét hướng của trang
  đầu tiên (hoặc trang đầu tiên trong danh sách trang được chọn) để áp dụng
  cho cả lệnh in đó — chưa xử lý được việc xoay theo từng trang riêng lẻ trong
  cùng 1 lệnh in.
- Gộp việc đặt số bản in và hướng giấy vào driver máy in thành 1 lần mở/đóng
  driver duy nhất (trước đó là 2 lần riêng biệt) — gọn hơn, giảm số lần truy
  cập driver không cần thiết.
- **Sửa lỗi quan trọng khác**: 1 số file PDF nhỏ vẫn bị "miss" (không xác nhận
  được lệnh in đã vào hàng đợi, dù thực chất máy in không nhận được) — do cơ
  chế xác nhận ở bản trước chỉ so sánh SỐ LƯỢNG job trong hàng đợi. Nếu job của
  file trước vừa in xong và rời hàng đợi ĐÚNG LÚC job của file hiện tại vừa
  được thêm vào, tổng số lượng không đổi khiến chương trình tưởng nhầm là chưa
  có gì thay đổi. Đã sửa bằng cách so sánh theo **Job ID cụ thể** (mỗi lệnh in
  luôn được Windows cấp 1 ID mới, không bao giờ trùng lặp) thay vì đếm số
  lượng — đáng tin cậy hơn hẳn, đặc biệt với các file nhỏ in rất nhanh.
- **Sửa lỗi "chọn in 1 trang nhưng in cả 2 trang"**: nguyên nhân là ô nhập
  trang khi chọn "Chỉ in trang chỉ định" có giá trị mặc định là "1,2" — nếu
  chỉ tích chọn radio mà quên sửa nội dung ô nhập, chương trình vẫn in đúng
  theo nội dung đang hiển thị (cả 2 trang). Đã đổi giá trị mặc định thành "1",
  đồng thời thêm dòng cảnh báo rõ ràng ngay trong màn cài đặt để nhắc kiểm tra
  lại trước khi bấm Tiếp tục.

## v3.0 (PyQt6) — Chuyển giao diện từ tkinter sang PyQt6
Viết lại toàn bộ phần giao diện bằng PyQt6, giữ nguyên 100% phần nghiệp vụ in ấn (Excel COM, xử lý PDF, xác nhận hàng đợi máy in, đặt cấu hình máy in...).
Kéo thả sắp xếp thứ tự file dùng cơ chế InternalMove có sẵn của QListWidget thay vì tự viết tay như bản tkinter.
Màn cài đặt dùng QScrollArea để cuộn được khi nội dung dài.

## v3.1 (PyQt6) — Hoàn thiện bản chuyển đổi PyQt6
Rà soát và dọn dẹp bản chuyển đổi PyQt6: xóa import thừa không dùng đến (QSizePolicy), thêm ghi chú giấy phép GPL v3 vào docstring đầu file.
Đã kiểm chứng: luồng nền (threading.Thread + queue.Queue + QTimer poll), cơ chế đóng cửa sổ an toàn khi đang in, và kéo thả sắp xếp danh sách file đều hoạt động đúng — không phát hiện bug chức năng nào.

## v3.2 (PyQt6) — In PDF trực tiếp qua GDI, không qua app ngoài
Thay đổi kiến trúc lớn: bỏ hoàn toàn cách in PDF cũ (gọi verb "print" của Windows để nhờ ứng dụng đọc PDF mặc định như Edge/Acrobat tự mở và tự in). Giờ chương trình tự render từng trang PDF thành ảnh (qua thư viện PyMuPDF) rồi gửi thẳng cho driver máy in qua GDI — không còn phụ thuộc bất kỳ ứng dụng đọc PDF nào của hệ thống.
Giải quyết dứt điểm 2 vấn đề cố hữu của cách cũ:
Miss cả file: cách cũ là lệnh "bắn rồi quên" (fire-and-forget) — nếu app đang bận xử lý file trước, yêu cầu in file tiếp theo có thể bị bỏ qua âm thầm. Cách mới dùng StartDoc/EndDoc đồng bộ, không có khái niệm "bắn rồi quên" nữa.
Miss trang: cách cũ phụ thuộc "trí nhớ" cấu hình (khoảng trang, số bản) riêng của từng app đọc PDF, ngoài tầm kiểm soát. Cách mới tự dựng toàn bộ cấu hình in (DEVMODE) cho mỗi lần in mà không đọc/ghi gì vào máy in mặc định của hệ thống — không có "trí nhớ" nào để bị dính.
Bỏ hẳn dependency pypdf, thay bằng PyMuPDF (fitz) và Pillow (PIL) để đọc/render PDF.
Bỏ hẳn toàn bộ cơ chế tạo/xóa file PDF tạm (không còn cần trích trang ra file tạm trung gian nữa) — đơn giản hóa đáng kể, giảm nguy cơ lỗi MAX_PATH và rác file tạm đã từng gặp ở các bản trước.
Không còn cần thay đổi cấu hình máy in MẶC ĐỊNH của hệ thống — loại bỏ hoàn toàn rủi ro ảnh hưởng tới các lần in tay khác của người dùng.

## v3.4 (PyQt6) — Thêm khung báo lỗi riêng, sửa lỗi file PDF 0 trang
Thêm khung báo lỗi riêng biệt ở màn tiến trình (nền đỏ nhạt, viền đỏ, tiêu đề "⚠ File/sheet bị lỗi (N):")
Sửa lỗi: file PDF 0 trang (rỗng hoặc hỏng) từng gây crash IndexError — giờ báo rõ "Lỗi: file PDF không có trang nào (có thể bị rỗng hoặc hỏng)".
Mở rộng nhận diện lỗi trong khung báo lỗi, bắt thêm trường hợp hiếm "Không có sheet nào được in" (trước đây chỉ bắt các dòng chứa chữ "lỗi").

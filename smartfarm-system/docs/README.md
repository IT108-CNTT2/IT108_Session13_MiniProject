# SmartFarm System — Quản lý nhật ký và thiết bị

> Bài thực hành **PowerShell + Pipeline + Excel TRIM/VLOOKUP + Markdown + mô hình IPO/DIKW**. Hệ thống tự động trích xuất sự cố từ nhật ký máy chủ, thống kê số lỗi và đối soát tên thiết bị theo mã linh kiện.

## 1. Mục đích và cấu trúc dự án

**Mục đích:** Tổ chức dữ liệu theo cấu trúc thư mục rõ ràng; lọc các log có mức `ERROR` hoặc `CRITICAL`; tạo báo cáo sự cố; làm sạch mã linh kiện và tra cứu tên thiết bị trong Excel; giải thích luồng xử lý theo IPO và DIKW.

```text
smartfarm-system/
├── docs/
│   └── README.md                # Tài liệu hướng dẫn & phân tích
├── scripts/
│   └── auto_report.ps1          # Script PowerShell tự động
├── data/
│   ├── log-he-thong.log         # Nhật ký máy chủ gốc (24 dòng)
│   └── quan-ly-thiet-bi.xlsx    # File Excel đã hoàn thành (2 sheet)
├── backup/
│   └── bao-cao-su-co.txt       # Báo cáo được tạo từ log
└── .env                        # File cấu hình ẩn (không có mật khẩu)
```

**Lưu ý:** Đây là các đường dẫn **tương đối**, không cố định vào ổ `C:\`, nên có thể chép cả thư mục dự án sang máy khác.

## 2. Checklist công việc

- [x] Tạo cây thư mục `smartfarm-system/{docs,scripts,data,backup}` theo yêu cầu.
- [x] Đặt hai file dữ liệu đầu vào vào `data/` và tạo `.env` cấu hình ẩn ở thư mục gốc.
- [x] Viết script `scripts/auto_report.ps1` dùng Pipeline, `Get-Content`, `Select-String`, `Write-Output`, `>` và `>>`.
- [x] Xuất sẵn `backup/bao-cao-su-co.txt` từ dữ liệu gốc; kiểm tra có **8 sự cố**.
- [x] Hoàn thiện Excel: `TRIM`, `VLOOKUP(...,FALSE)`, khóa vùng `$`, `AutoFilter`, không có `#N/A`.
- [x] Kiểm tra thủ công script không có lệnh xóa hàng loạt và ghi chú `# AUDIT`.
- [ ] Thực thi lại script trên **Windows PowerShell** của máy nộp bài để xác nhận môi trường.

## 3. Hướng dẫn chạy PowerShell

Yêu cầu Windows PowerShell 5.1 hoặc PowerShell 7+, có quyền đọc thư mục `data` và ghi vào `backup`.

**Mở PowerShell tại thư mục `smartfarm-system`, sau đó chạy:**

```powershell
cd .\scripts
.\auto_report.ps1
```

Nếu Windows chặn thực thi vì `ExecutionPolicy`, chỉ cho phép **lần chạy cụ thể** (không thay đổi chính sách toàn máy):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\auto_report.ps1
```

Có thể chạy trực tiếp từ thư mục gốc mà không cần đổi vị trí hiện tại:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\auto_report.ps1
```

**Cách tạo thư mục bằng PowerShell nếu tự thực hành từ đầu:**

```powershell
New-Item -ItemType Directory -Path .\smartfarm-system
New-Item -ItemType Directory -Path .\smartfarm-system\docs
New-Item -ItemType Directory -Path .\smartfarm-system\scripts
New-Item -ItemType Directory -Path .\smartfarm-system\data
New-Item -ItemType Directory -Path .\smartfarm-system\backup
```

**Đặt thuộc tính Hidden cho `.env` khi tự thao tác tại thư mục gốc:**

```powershell
attrib +h .env
Get-Item .\.env -Force | Select-Object Name, Attributes
```

Script cũng tự đặt thuộc tính `Hidden` cho `.env` khi chạy trên Windows.

**Kiểm tra báo cáo sau khi chạy:**

```powershell
Get-Content ..\backup\bao-cao-su-co.txt
```

> Lưu ý: Dòng cuối `Get-Content` chạy trong thư mục `scripts`. Khi chạy từ gốc, dùng `Get-Content .\backup\bao-cao-su-co.txt`.

### Giải thích script

- `Split-Path -Parent $PSScriptRoot` tìm thư mục gốc theo **vị trí script** (không phụ thuộc đang đứng ở đâu trong PowerShell).
- Phần `Get-Item .env -Force` kết hợp thuộc tính `Hidden` đặt file `.env` ở chế độ **ẩn trên Windows** mỗi khi chạy script. Không chứa thông tin bí mật trong file `.env` mẫu.
- `Join-Path` ghép đường dẫn tương đối an toàn trên máy Windows.
- `Write-Output '...' > $reportFile` **tạo/ghi mới** tiêu đề báo cáo, đảm bảo mỗi lần chạy chỉ có một bản báo cáo cập nhật.
- `Get-Content | Select-String -Pattern '\[(ERROR|CRITICAL)\]'` đọc từng dòng rồi **lọc chính xác** hai mức độ cần xử lý.
- `ForEach-Object { $_.Line } >> $reportFile` **nối thêm** nội dung dòng gốc vào báo cáo (không làm mất tiêu đề).
- `Measure-Object | Select-Object -ExpandProperty Count` đếm số sự cố, sau đó `Write-Output ... >>` ghi dòng `TONG SO SU CO: 8`.
- `# AUDIT` xác nhận đã rà soát: không có `Remove-Item -Recurse -Force`, `del /f /s /q` hay lệnh phá hủy dữ liệu hàng loạt.

## 4. Báo cáo sự cố: kết quả từ log thật

Nguồn: `data/log-he-thong.log` (ngày ghi log: **12/10/2026**, tổng cộng 24 dòng).

| Chỉ số | Số lượng |
|:--|--:|
| Tổng số dòng log | 24 |
| Sự cố `ERROR` | 6 |
| Sự cố `CRITICAL` | 2 |
| **Tổng số sự cố** (`ERROR` + `CRITICAL`) | **8** |
| Sự cố nghiêm trọng liên quan `sensor-03` | 4 |

Các dòng liên quan `sensor-03`: **08:15:01**, **09:03:17**, **10:02:48** (đều `Connection TIMEOUT after 30s`) và **11:12:37** (`Power supply failure`). Các dòng `WARNING` và `INFO` không nằm trong báo cáo vì đề chỉ yêu cầu lọc `ERROR/CRITICAL`.

Báo cáo mẫu đã được tạo sẵn tại `backup/bao-cao-su-co.txt`. Chạy `auto_report.ps1` sẽ cập nhật lại thời điểm lập báo cáo theo giờ máy thực thi.

## 5. Excel — Đối soát tài nguyên (TRIM + VLOOKUP)

File: `data/quan-ly-thiet-bi.xlsx` gồm hai sheet, giữ nguyên dữ liệu gốc.

### 5.1. Sheet `danh-muc-linh-kien`

Danh mục có 6 mã linh kiện. Bảng sau lấy từ file Excel đầu vào:

| Mã linh kiện | Tên thiết bị | Đơn giá (VND) | Loại storage |
|:--|:--|--:|:--|
| LK01 | Raspberry Pi 4 Model B | 1.450.000 | SSD |
| LK02 | Intel NUC i3 | 6.200.000 | SSD |
| LK03 | Orange Pi 5 | 2.300.000 | SSD |
| LK04 | Dell Optiplex Micro | 5.400.000 | HDD |
| LK05 | Mini PC Fanless N100 | 4.200.000 | SSD |
| LK06 | Synology DS220j NAS | 3.900.000 | HDD |

### 5.2. Sheet `trang-thai-tram`

Sheet có 10 trạm, 4 cột gốc (`ma_tram`, `ma_linh_kien`, `dung_luong_ram`, `ngay_lap_dat`). Đã bổ sung hai cột mới:

- **E — `ma_linh_kien_sach`**: tại ô `E2`, công thức `=TRIM(B2)`; kéo xuống `E11`. Ví dụ `" LK01" → "LK01"`, `"LK05 " → "LK05"`.
- **F — `ten_thiet_bi`**: tại ô `F2`, công thức dưới đây; kéo xuống `F11`. `FALSE` yêu cầu tra cứu **chính xác**, `$A$2:$D$7` khóa tuyệt đối vùng danh mục, `IFERROR` hiển thị thông báo dễ hiểu khi mã mới không tồn tại:

```excel
=IFERROR(VLOOKUP(E2,'danh-muc-linh-kien'!$A$2:$D$7,2,FALSE),"Không tìm thấy")
```

> Excel tại một số máy dùng dấu chấm phẩy (`;`) thay dấu phẩy (`,`), tùy thiết lập vùng. Trong file đính kèm, công thức được ghi theo chuẩn `.xlsx` và đã cho ra đúng tên thiết bị cho cả 10 trạm.

**Ví dụ dữ liệu đối soát:**

| Trạm | Mã linh kiện gốc | Mã sau TRIM | Tên thiết bị sau VLOOKUP |
|:--|:--|:--|:--|
| ST01 | `LK02` | `LK02` | Intel NUC i3 |
| ST02 | ` LK01` | `LK01` | Raspberry Pi 4 Model B |
| ST03 | `LK05 ` | `LK05` | Mini PC Fanless N100 |
| ST05 | `LK02 ` | `LK02` | Intel NUC i3 |
| ST08 | ` LK04` | `LK04` | Dell Optiplex Micro |

**Kiểm tra đạt:** 10/10 dòng tra cứu thành công, không có `#N/A`. Vùng bảng đã bật **AutoFilter** trên hàng tiêu đề bằng Excel Table, không chèn hàng trống vào vùng dữ liệu và không dùng Merge Cell.

## 6. Phân tích mô hình IPO — Input, Process, Output

| Thành phần | Ứng dụng trong SmartFarm |
|:--|:--|
| **Input (Đầu vào)** | `log-he-thong.log` chứa timestamp, mức độ, thiết bị, mô tả; `quan-ly-thiet-bi.xlsx` chứa bảng danh mục linh kiện và tình trạng trạm. |
| **Process (Xử lý)** | PowerShell đọc log, dùng Pipeline + `Select-String` lọc `ERROR/CRITICAL`, đếm bằng `Measure-Object`; Excel chuẩn hóa mã bằng `TRIM` và đối chiếu qua `VLOOKUP` khớp chính xác. |
| **Output (Đầu ra)** | `backup/bao-cao-su-co.txt` có 8 sự cố và thời điểm xuất; file Excel có hai cột mã sạch và tên thiết bị, kèm bộ lọc tự động. |

## 7. Phân tích DIKW — Data, Information, Knowledge, Wisdom

**DIKW** mô tả cách chuyển dữ liệu thô thành thông tin, hiểu biết và quyết định thực tế.

| Phát biểu trong đề | Tầng DIKW | Lý do |
|:--|:--|:--|
| (a) Quyết định thay modem tại trạm `sensor-03` trước thứ Hai. | **Wisdom (Trí tuệ / quyết định)** | Hành động dựa trên phân tích. **Tuy nhiên**, cần kiểm tra nguồn điện trước khi quyết định thay modem: log có lỗi `Power supply failure`, chưa chứng minh modem hỏng. |
| (b) `2026-10-12 08:15:01 [ERROR] sensor-03 Connection TIMEOUT after 30s` | **Data (Dữ liệu)** | Một dòng ghi nhận nguyên gốc, chưa tổng hợp. |
| (c) Hệ thống ghi nhận 8 sự cố, trong đó 4 sự cố đến từ `sensor-03`. | **Information (Thông tin)** | Số liệu đã được lọc, tổng hợp theo sự cố và thiết bị. Con số đúng với log đầu vào. |
| (d) `sensor-03` liên tục gặp lỗi kết nối; khả năng vấn đề liên quan nguồn điện, cần kiểm tra nguyên nhân. | **Knowledge (Tri thức)** | Suy luận từ 3 lỗi timeout và 1 lỗi nguồn. Đây là **giả thuyết có căn cứ**, không phải kết luận chắc chắn rằng mọi timeout đều do nguồn điện. |

**Hướng xử lý khoa học hơn:** ưu tiên đo kiểm nguồn cấp và dây nguồn của `sensor-03`, sau đó kiểm tra đường truyền, kết nối và modem nếu lỗi còn tiếp diễn. Việc thay modem chỉ nên quyết định sau khi có chứng cứ.

## 8. Bẫy phân tích: tốc độ truyền file 500 MB qua mạng 100 Mbps

- Kích thước tệp: **500 MB** (megabyte).
- Tốc độ đường truyền: **100 Mbps** (megabit mỗi giây).
- Đổi sang MB/s: `100 Mbps / 8 = 12,5 MB/s` do **1 byte = 8 bit**.
- Thời gian lý thuyết: `500 MB / 12,5 MB/s = 40 giây`.

**Kết luận: 40 giây trong điều kiện lý tưởng.** Thực tế có thể lâu hơn do overhead giao thức, tốc độ đĩa, chất lượng mạng và cách tính MB/MiB.

## 9. Tự kiểm tra trước khi nộp

1. Mở ZIP, bảo đảm có bốn thư mục `docs`, `scripts`, `data`, `backup` và file `.env` ẩn ở thư mục gốc.
2. Xác nhận `scripts/auto_report.ps1` chứa comment `# AUDIT`, sử dụng `Get-Content | Select-String` và `>` / `>>`.
3. Từ thư mục `scripts`, chạy PowerShell theo hướng dẫn; file báo cáo phải xuất hiện ở `../backup/bao-cao-su-co.txt` với `TONG SO SU CO: 8`.
4. Mở Excel, kiểm tra công thức tại `E2:E11` và `F2:F11`, tra cứu ra tên đúng và có biểu tượng bộ lọc trên hàng tiêu đề.
5. Mở `docs/README.md` ở GitHub hoặc VS Code Markdown Preview để xem đúng Heading, Code Block, Checklist, bảng và phân tích DIKW.

---

*Tệp được chuẩn bị từ 2 nguồn dữ liệu thực tế kèm bài; không tự bổ sung các sự cố hoặc linh kiện không tồn tại trong dữ liệu.*

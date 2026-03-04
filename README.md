# TextInputter

Ứng dụng Windows (C# WinForms) hỗ trợ nhân viên giao hàng nhập liệu hóa đơn nhanh từ ảnh chụp, xuất thẳng vào file Excel theo định dạng của khách.

---

## Tính năng chính

### 🔍 Tab OCR — Xử lý hàng loạt
- Chọn thư mục ảnh → quét toàn bộ bằng **Google Cloud Vision**
- Tự động extract 12 fields: `SHOP`, `TÊN KH`, `MÃ`, `ĐỊA CHỈ`, `QUẬN`, `PHƯỜNG`, `TÊN ĐƯỜNG`, `TIỀN THU`, `TIỀN SHIP`, `NGÀY LẤY`, `NGƯỜI ĐI`, `GHI CHÚ`
- **Gemini fallback tự động** khi regex thiếu field — thử tuần tự 5 model (2.5-flash-lite → 2.5-pro)
- Tự động tra **phí ship 4 cấp** theo địa chỉ:
  - Tier 3: theo phường cụ thể (Q8 chia đôi P.5–16)
  - Tier 2.8: theo tên đường cụ thể (Yên Thế, Vinhome, Đặng Nguyên Cẩn...)
  - Tier 2.5: phường → quận (qua bảng map)
  - Tier 2: theo quận
- Tự động điền **người đi** theo khu vực (c.hieu / c.cuong / a.quyen / An Tam dd-MM)
- Hỗ trợ **manual override** người đi/người lấy
- **Xuất Excel** → user chọn file đích qua dialog → ghi vào sheet `dd-MM`
- **Tính tiền** riêng: tính TIỀN HÀNG + sinh bảng tổng per SHOP và per NGƯỜI ĐI

### 📋 Tab Invoice — Xem & tính báo cáo ngày
- Mở file Excel bất kỳ, tự detect header
- Tính tổng TIỀN THU / TIỀN SHIP / TIỀN HÀNG per SHOP
- Lưu báo cáo ngày ra `DailyTotalReport.xlsx`

### ✍️ Tab Manual Input
- Nhập tay thông tin đơn hàng (đang phát triển)

---

## Cài đặt & chạy

### Yêu cầu
- Windows 10/11
- .NET 8.0 SDK (để build) hoặc chỉ cần Runtime (để chạy bản publish)

### Build & chạy từ source
```powershell
git clone <repo>
cd TextInputter
dotnet run
```

### Publish ra file .exe standalone
```powershell
dotnet publish -c Release -r win-x64 --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -o publish```
File .exe xuất ra trong thư mục `publish\`.

---

## Cấu hình

### Google Cloud Vision (OCR chính)
1. Tạo Google Cloud project, bật **Vision API**
2. Tạo Service Account → tải JSON key
3. Đặt file JSON **cạnh file .exe** (hoặc cạnh thư mục project khi dev)
4. Sửa tên file trong `AppConstants.GOOGLE_CREDENTIAL_FILE`

### Gemini API (OCR fallback)
1. Lấy API key tại [aistudio.google.com](https://aistudio.google.com/apikey) (free tier đủ dùng)
2. Điền vào `AppConstants.GEMINI_API_KEY`

> ⚠️ **Không commit API key / credential JSON lên git public**

---

## Bảng phí ship hiện tại

| Quận/Khu | Phí | Ghi chú |
|----------|-----|---------|
| Q1, Q3 | 20k | |
| Q4, Q5, Q10, Q11 | 25k | |
| Q6 | 25k | đường Đặng Nguyên Cẩn → 30k |
| Q7, Q2, Q12 | 30k | |
| Q8 | 25k | P.5–7, P.11–16 → 30k |
| Q9 | 30k | Vinhome Grand Park → 35k |
| Bình Thạnh, Phú Nhuận | 20k | |
| Gò Vấp, Tân Phú | 25k | |
| Tân Bình | 25k | đường Yên Thế, Quách Văn Tuấn → 30k |
| Bình Tân, Thủ Đức | 30k | |
| Bình Chánh, Hóc Môn, Nhà Bè | 35k | |
| Củ Chi | 40k | |
| Cần Giờ | 50k | |

Sửa bảng phí: `main/AppConstants.cs` → `SHIPPING_FEES_BY_QUAN` / `SHIPPING_FEES_BY_WARD` / `SHIPPING_FEES_BY_STREET`

---

## Cấu trúc project

```
main/
├── AppConstants.cs          ← tất cả constants/hardcoded values
├── MainForm.cs              ← shared fields + constructor
├── tabs/
│   ├── OcrTab.cs            ← OCR batch logic
│   ├── InvoiceTab.cs        ← báo cáo ngày
│   └── ManualInputTab.cs    ← nhập tay (WIP)
└── Services/
    ├── OCRTextParsingService.cs   ← parse OCR text → fields
    ├── GeminiService.cs           ← Gemini fallback
    ├── ExcelInvoiceService.cs     ← ghi/cập nhật Excel
    ├── OCRInvoiceMapper.cs        ← tra ship fee, người đi
    └── AddressParser.cs           ← tách địa chỉ VN
```

Chi tiết kỹ thuật xem [`ARCHITECTURE.md`](ARCHITECTURE.md).

---

## Workflow thực tế

```
1. Chụp ảnh hóa đơn → chép vào thư mục data/
2. Tab OCR → Chọn Thư Mục → Bắt Đầu
3. Kiểm tra log kết quả (txtProcessLog)
4. Xuất Excel → chọn file Excel đích → OK
5. Tính Tiền → sinh bảng tổng tự động
6. Mở Excel kiểm tra kết quả
```

---

## Hành vi đặc biệt

- **Hàng sỉ / ship gộp** (không có MÃ): TÌNH TRẠNG = `"hàng sỉ"`, không tô đỏ cột MÃ
- **Người đi "An Tam"**: tự động append ngày `dd-MM` vào cuối (vd: `"An Tam 15-03"`)
- **Duplicate MÃ**: nếu mã đơn đã tồn tại trong bất kỳ sheet nào → **cập nhật** thay vì thêm mới
- **TIỀN THU** trong Excel = Tổng thanh toán OCR + TIỀN SHIP (để tính đúng doanh thu)
- **TIỀN HÀNG** = TIỀN THU - TIỀN SHIP (tính sau khi bấm "Tính Tiền")

# TextInputter — Architecture Guide

> **Mục đích:** Tài liệu này giúp người mới vào dự án hiểu cấu trúc code, flow hoạt động, và biết phải thêm/sửa code ở đâu khi cần.

---

## 1. Tech Stack

| Layer | Technology |
|---|---|
| Runtime | .NET 8.0 Windows, C# |
| UI | Windows Forms (WinForms) |
| Excel I/O | [ClosedXML 0.102.3](https://github.com/ClosedXML/ClosedXML) |
| OCR | [Google Cloud Vision V1 3.8.0](https://cloud.google.com/vision) |
| Credentials | `textinputter-4a7bda4ef67a.json` (Google Service Account) |

---

## 2. Cấu trúc thư mục

```
TextInputter/
├── TextInputter.csproj          # Project file (packages, targets)
├── ARCHITECTURE.md              # (file này)
├── README.md
│
├── main/                        # Toàn bộ source code
│   ├── Program.cs               # Entry point — chỉ gọi Application.Run(new MainForm())
│   ├── AppConstants.cs          # TẤT CẢ hardcoded values tập trung tại đây ← đọc khi cần config
│   │
│   ├── MainForm.cs              # Shared fields + constructor + shared helpers
│   ├── MainForm.Designer.cs     # Form-level skeleton: panelTop/Left/Bottom, tabMainControl + 4 TabPages
│   │                            # (KHÔNG chứa tab-specific controls — đã tách sang InvoiceTab.UI.cs)
│   │
│   ├── tabs/                    # Partial classes của MainForm — mỗi tab 1–2 file
│   │   │
│   │   ├── OcrTab.cs            # OCR tab: InitializeOCRTab() + toàn bộ logic + control fields
│   │   │                        #   Controls: txtRawOCRLog, txtProcessLog, txtNguoiDiOCR, txtNguoiLayOCR
│   │   │
│   │   ├── InvoiceTab.cs        # Invoice logic: BtnOpenExcel, Calculate, DailyReport, Save
│   │   ├── InvoiceTab.UI.cs     # Invoice UI: InitializeInvoiceTabUI() + control declarations
│   │   │                        #   Controls: tabExcelSheets, panelExcelButtons, dgvInvoice, lblInvoiceTotal, ...
│   │   │
│   │   └── ManualInputTab.cs    # Manual Input: InitializeManualInputTab() + logic (UI inline, không cần file riêng)
│   │
│   ├── Services/                # Business logic (không phụ thuộc UI)
│   │   ├── OCRTextParsingService.cs   # Parse raw OCR text → extract 12 fields
│   │   ├── ExcelInvoiceService.cs     # Ghi dữ liệu invoice vào file Excel của khách
│   │   ├── OCRInvoiceMapper.cs        # Model OCRInvoiceData + helper mapping (ít dùng)
│   │   └── AddressParser.cs           # Parse địa chỉ VN → SoNha, TenDuong, Phuong, Quan
│   │
│   └── utils/
│       ├── UIHelper.cs               # Factory methods tạo WinForms controls + RichTextBox search
│       └── AddressParsingDialog.cs   # Dialog xác nhận địa chỉ đã parse
│
├── resources/
│   └── app.ico
│
└── data/
    └── sample/                  # File mẫu để test
        ├── excel/
        └── images/
```

---

## 3. Pattern: Partial Classes (quan trọng!)

`MainForm` được **split thành nhiều file** bằng cơ chế `partial class` của C#:

```
MainForm.cs              → shared fields, constructor, shared helpers
MainForm.Designer.cs     → form-level skeleton (panelTop/Left/Bottom, tabMainControl + TabPages)

tabs/OcrTab.cs           → partial class MainForm { control fields + InitializeOCRTab() + logic }
tabs/InvoiceTab.UI.cs    → partial class MainForm { control fields + InitializeInvoiceTabUI() }
tabs/InvoiceTab.cs       → partial class MainForm { logic handlers }
tabs/ManualInputTab.cs   → partial class MainForm { InitializeManualInputTab() + logic }
```

**Quy tắc phân tách UI / Logic:**

| File | Chứa gì |
|---|---|
| `MainForm.Designer.cs` | Chỉ form-level: panelTop, panelLeft, panelBottom, tabMainControl, 4 TabPage |
| `<Tab>.UI.cs` | Control field declarations + `Initialize<Tab>UI()` — chỉ layout, không có event logic |
| `<Tab>.cs` | Event handlers, business logic, service calls |
| `MainForm.cs` | Shared fields, constructor (gọi cả `Initialize...UI()` + `Initialize...Tab()`), shared helpers |

**Ý nghĩa thực tế:**
- Tất cả đều **cùng 1 class** — mọi field/method ở file nào cũng truy cập được từ file khác.
- Khi thêm tab mới → tạo `tabs/NewTab.UI.cs` (controls) + `tabs/NewTab.cs` (logic), gọi `InitializeNewTabUI()` trong constructor.
- Khi thêm shared helper → viết vào `MainForm.cs`.
- **Không được** đặt control declarations hay `InitializeComponent()` calls vào tab logic files.

---

## 4. Shared Fields (MainForm.cs)

| Field | Type | Mô tả |
|---|---|---|
| `folderPath` | `string` | Đường dẫn folder ảnh OCR đang chọn |
| `imageFiles` | `List<string>` | Danh sách file ảnh trong folder |
| `isProcessing` | `bool` | Flag chống double-click khi đang xử lý |
| `visionClient` | `ImageAnnotatorClient` | Google Vision client (init trong `InitializeServices`) |
| `_ocrParsingService` | `OCRTextParsingService` | Service parse OCR text |
| `mappedDataList` | `List<Dictionary<string,string>>` | Cache kết quả OCR đã map (dùng để export) |

> `_excelInvoiceService` đã bị xóa — `ExcelInvoiceService` được khởi tạo nhưng chưa wire vào UI nên loại bỏ tránh nhầm lẫn.

---

## 5. Flow chính

### 5A. Excel Viewer + Daily Report (InvoiceTab.cs)

```
User click 📁 Mở File
    └─ BtnOpenExcel_Click()
         └─ LoadExcelFile()              ← đọc Excel bằng ClosedXML
              └─ DetectHeaderRow()       ← tìm header dựa vào HEADER_ROW_KEYWORDS
                   └─ MapColumnIndices() ← gán cột SHOP, TIỀN THU, TIỀN SHIP, ...

User click 🧮 Tính
    └─ BtnCalculateExcelData_Click()
         └─ CalculateAllRows()          ← vòng lặp qua tất cả row, tính tổng
              └─ DisplayDailyReport()   ← hiện bảng tổng cuối màn hình

User click 💾 Lưu
    └─ SaveDailyReportToExcel()         ← ghi DailyTotalReport.xlsx
```

### 5B. OCR Batch (OcrTab.cs)

```
User click 📁 Chọn Thư Mục
    └─ SelectOCRFolder()
         └─ GetImageFiles()             ← lọc .jpg/.png/.webp (MainForm.cs)

User nhập Người Đi / Người Lấy (TextBox trong UI)

User click ▶ Bắt Đầu
    └─ btnStart_Click() → ProcessImages() [async]   ← vòng lặp qua ảnh đã chọn
              ├─ CallGoogleVisionOCR()  ← gửi ảnh lên Google Vision (MainForm.cs)
              ├─ CleanOCRText()         ← lọc garbage lines (MainForm.cs)
              ├─ _ocrParsingService.ExtractAllFields()   ← parse 10 fields
              ├─ inject NGƯỜI ĐI / NGƯỜI LẤY từ UI
              ├─ OCRInvoiceMapper.GetShipFeeByQuan()     ← auto-fill TIỀN SHIP theo quận
              ├─ → append vào mappedDataList
              └─ → ghi raw OCR vào txtRawOCRLog, kết quả map vào txtProcessLog

User click � Xuất Excel
    └─ ExportMappedDataToExcel()        ← user chọn file Excel, ghi vào sheet dd-MM
```

### 5C. Manual Input (ManualInputTab.cs)

```
User điền 17 fields vào form
    └─ SaveManualEntry()
         └─ ⚠️ TODO: hiện chỉ MessageBox — chưa ghi vào Excel
```

---

## 6. Services

### `OCRTextParsingService`
**Input:** raw OCR text (string từ Google Vision)  
**Output:** `Dictionary<string, string>` chứa các fields, + `List<string>` các fields bị thiếu

| Method | Mô tả |
|---|---|
| `ExtractAllFields(text, out fields)` | Public entry point — extract 10 fields (NGƯỜI ĐI/LẤY do UI cung cấp, TIỀN SHIP không còn required) |
| `ExtractAddressLine(text)` | Private — lấy dòng "địa chỉ:" **cuối cùng** hợp lệ (bỏ qua địa chỉ shop CN1/CN2). Match: `"địa chỉ"`, `"địa chi"` (OCR drop dấu), `"dia chi"`, `"address"` |
| `ExtractAmountLine(text, keywords)` | Private — tìm số tiền theo từ khoá; xử lý cả số cùng dòng lẫn số ở dòng tiếp theo |
| `NormalizeToThousands(raw)` | Private — chuẩn hóa về nghìn đồng (1,500,000 → 1500) |
| `ExtractDate(text)` | Private — parse ngày từ text |

**Edge cases đã xử lý (từ data thật):**

| Input thực tế | Vấn đề | Cách xử lý |
|---|---|---|
| `Địa Chi: 132 bên Vân đồn,p6,q4 - -` | OCR drop dấu `ỉ` → `"chi"` thay vì `"chỉ"` | Match thêm `"địa chi"` (có dấu `ị`) + `"dia chi"` (không dấu) |
| Hóa đơn có 2 dòng `Địa Chi/Chỉ:` (shop CN1 + khách hàng) | Parse nhầm địa chỉ shop | Lấy dòng **cuối cùng** hợp lệ; bỏ qua nếu chứa `CN\d / HOTLINE / SĐT` |
| `132 bên Vân đồn,p6,q4 - -` | Trailing garbage `- -` | Strip `[\s\-]+$` sau khi extract |
| `A25 hotel ( phòng 706) 184 nguyễn trãi, phường phạm ngũ lão, q1` | Số nhà phức tạp (tên khách sạn + số phòng + số nhà) | `ExtractHouseAndStreet` dùng greedy regex lấy đến số cuối cùng |
| `So HD: HD130781` (không dấu) | OCR drop dấu `ố` → `"So"` | Regex `So\s*H[ĐD]` đã cover |
| Số tiền trên dòng riêng (`Tổng tiền hàng:\n1,500,000`) | Số không cùng dòng keyword | `ExtractAmountLine` check thêm `lines[i+1]` |
| `TIỀN SHIP` không có trên hóa đơn | Field trống → lỗi validation | Không còn required — auto-fill từ bảng phí theo quận |
| `363-365-367, 363 Đ. Hùng Vương - Khải Nam Transpost – –` | Số nhà là dãy số có `-`, tên business rác sau ` - ` | Strip ` - <tên không phải địa chỉ>` ở cuối; `Đ.` không bị strip vì được loại trừ khỏi regex |
| `Địa chỉ: 11 In Dung Vương Phường An Đông TP HCM ạ` | `"Phường An Đông"` và `"TP HCM"` cuối địa chỉ | Strip `Phường <tên>` + `TP HCM` ở cuối trước khi pass vào AddressParser |

### `ExcelInvoiceService`
**Mục đích:** Ghi dữ liệu OCR vào file Excel của khách (20 cột cố định)  
**File Excel:** hardcoded `"CHÂU NGÂN- THÁNG 2.2026- ĐỐI SOÁT.xlsx"` ⚠️  
**Trạng thái:** ⚠️ Chưa được wire vào UI — `ExportMappedDataToExcel()` trong `OcrTab.cs` vẫn dùng ClosedXML trực tiếp.

| Method | Mô tả |
|---|---|
| `InvoiceExists(ma)` | Kiểm tra mã đơn đã tồn tại trong sheet chưa |
| `ExportInvoice(data, sheetName)` | Ghi 1 row vào sheet (tạo sheet nếu chưa có) |
| `GetAllInvoiceNumbers()` | Trả về tất cả mã đơn đã ghi |

### `AddressParser`
**Input:** string địa chỉ thô  
**Output:** `ParsedAddress { SoNha, TenDuong, Phuong, Quan, Confidence }`  
Có dictionary nội bộ cho quận/huyện TP.HCM + Hà Nội. **Phường không ảnh hưởng đến tính toán tiền ship.**

**Edge cases đã xử lý (từ data thật):**

| Input thực tế | Vấn đề | Cách xử lý |
|---|---|---|
| `5/1 phùng văn cung p2 phủ nhuận` | Không có dấu phẩy giữa các thành phần | Tự chèn phẩy trước `p<số>`, `q<số>`, `phường`, `quận` inline |
| `11 In Dung Vương` | Số nhà `11` bị nhận nhầm là Quận 11 | Bare number (`^\d{1,2}$`) chỉ match quận khi **toàn segment là số đó**; bỏ qua nếu segment có nhiều từ |
| `363-365-367, 363 Đ. Hùng Vương` | Dãy số nhà nhiều giá trị, tên đường ở segment kế | Khi `firstSeg` chỉ toàn số và `-` → dùng segment kế làm nguồn tên đường |
| `363 Đ. Hùng Vương` | `Đ.` (viết tắt Đường) trước tên đường | Regex riêng bắt `<số> Đ. <tên>` → SoNha + TenDuong |
| `phủ nhuận` / `phú nhuật` (OCR sai dấu) | Không khớp exact với `"phú nhuận"` | Fuzzy lookup: xóa dấu → match `"phu nhuan"` trong `DistrictNoDiacDict` |

### `OCRInvoiceMapper`
**Mục đích hiện tại:** Chứa model `OCRInvoiceData` và helper tra phí ship.  
> `MapToExcelColumns` và `ParseAndVerifyAddress` đã bị xóa (không có caller).

| Method / Class | Mô tả |
|---|---|
| `OCRInvoiceData` | Model class chứa tất cả fields của 1 invoice. Dùng bởi `ExcelInvoiceService` |
| `GetShipFeeByQuan(quan)` | Tra bảng `AppConstants.SHIPPING_FEES_BY_QUAN` theo quận, tự normalize không dấu. Trả `null` nếu không tìm thấy |

### `UIHelper`
Factory methods tạo controls đồng bộ style:
- `CreateLabelTextBox(label, width)` — tạo Label + TextBox ghép đôi
- `CreateButton(text, color)` — tạo Button với style chuẩn
- `CreateReadOnlyTextBox()` — TextBox read-only
- `CreateSectionLabel(text)` — Label tiêu đề section
- `CreateRichTextBoxSearchBar(parent, y, getTarget)` — tạo search bar (🔍 TextBox + ▼▲✕ + label X/Y) gắn vào một RichTextBox
- `SearchInRichTextBox(rtb, term, forward, idxHolder, lblResult)` — tìm kiếm, highlight vàng/cam, scroll đến match
- `ClearRichTextBoxHighlights(rtb)` — xóa toàn bộ highlight trong RichTextBox

---

## 7. Thêm tính năng mới — làm ở đâu?

| Muốn làm gì | File cần edit |
|---|---|
| Thêm tab mới | Tạo `tabs/NewTab.cs` với `partial class MainForm` |
| Thêm field mới vào OCR output | `OCRTextParsingService.ExtractAllFields()` |
| Thêm cột mới vào Excel export | `ExcelInvoiceService` + `OCRInvoiceData` |
| Thêm config/constant (data thuần) | `AppConstants.cs` |
| Thêm logic map/lookup OCR | `Services/OCRInvoiceMapper.cs` |
| Cập nhật bảng phí ship theo quận | `AppConstants.SHIPPING_FEES_BY_QUAN` |
| Thêm shared UI control style | `utils/UIHelper.cs` |
| Thêm search bar cho RichTextBox | `UIHelper.CreateRichTextBoxSearchBar()` |
| Thêm shared helper (dùng nhiều tab) | `MainForm.cs` |
| Thay đổi logic tính toán Excel Viewer | `InvoiceTab.cs` — `CalculateAllRows()` |
| Thay đổi cách detect header row | `InvoiceTab.cs` — `DetectHeaderRow()` + `AppConstants.HEADER_ROW_KEYWORDS` |
| Thay đổi cách OCR gọi Google | `MainForm.cs` — `CallPythonOCR()` |
| Thêm loại ảnh được chấp nhận | `MainForm.cs` — `GetImageFiles()` |

---

## 8. Danh sách Hardcoded — cần discuss để cải thiện

> Tất cả constant đã tập trung trong `AppConstants.cs`. Danh sách bên dưới là các mục **còn nằm rải rác** hoặc **cần input từ user thay vì code cứng**.

| # | Vị trí | Giá trị cứng | Vấn đề |
|---|---|---|---|
| 1 | `ExcelInvoiceService.cs` constructor | `"CHÂU NGÂN- THÁNG 2.2026- ĐỐI SOÁT.xlsx"` | Tên file client-specific, đổi tháng là lỗi |
| 2 | `AppConstants.PHI_SHIP_MOI_DON` | `5m` (5đ/đơn) | Business rule, nên cho user input |
| 3 | `AppConstants.COL_SODON_FALLBACK_IDX` | `17` | Phụ thuộc column index Excel cụ thể |
| 4 | `AppConstants.HEADER_ROW_KEYWORDS` | `{"SHOP", "Tình trạng"}` | Phụ thuộc template Excel của khách |
| 5 | `OcrTab.ExportMappedDataToExcel()` | 20-column header array | Client-specific Excel template |
| 6 | `AppConstants.DATE_FORMAT_SHEET` | `"dd-MM-yyyy"` | Sheet naming convention cứng |
| 7 | `OCRTextParsingService` | Tất cả regex keyword | Phụ thuộc format hóa đơn hiện tại |
| 8 | `AddressParser` | `DistrictDict`, `WardDict` | Chỉ cover TP.HCM |
| 9 | `AppConstants.GOOGLE_CREDENTIAL_FILE` | `"textinputter-4a7bda4ef67a.json"` | Credential file cứng cạnh .exe |
| 10 | `AppConstants.SHIPPING_FEES_BY_QUAN` | Bảng phí ship theo quận | Phụ thuộc hợp đồng vận chuyển hiện tại, chỉ cover TP.HCM |

**Hướng cải thiện đề xuất (discuss sau):**
- Item 1: Dùng `OpenFileDialog` để user chọn file Excel đích khi start, hoặc đọc từ `appsettings.json`
- Item 2, 3: Thêm "Settings" tab hoặc `config.json`
- Item 4, 5: Tách thành template config riêng theo khách hàng
- Item 9: Dùng environment variable hoặc `appsettings.json`

---

## 9. Các điểm cần hoàn thiện (TODO)

| File | Vị trí | Vấn đề |
|---|---|---|
| `ManualInputTab.cs` | `SaveManualEntry()` | Chưa ghi vào Excel — hiện chỉ hiện MessageBox |
| `MainForm.cs` | `CallPythonOCR()` | Tên hàm misleading (không call Python) — là Google Vision API |
| `OcrTab.cs` | `ExportMappedDataToExcel()` | Dùng ClosedXML thẳng với `Dictionary<string,string>` — chưa dùng `ExcelInvoiceService` |
| `ExcelInvoiceService.cs` | constructor | Tên file Excel hardcoded theo tháng — cần đổi mỗi tháng |

---

## 10. Warnings hiện tại (không block build)

| Warning | Nguồn | Giải thích |
|---|---|---|
| `CS8669` (×6) | `MainForm.Designer.cs` | Nullable annotation trong auto-generated code — bỏ qua |
| `CS0618` | `MainForm.cs:57` | `GoogleCredential.FromFile()` deprecated — vẫn hoạt động, fix sau |

---

## 11. Command để build file .exe

``` 
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o publish\ 
```


# TextInputter — Architecture Guide

> **Mục đích:** Tài liệu này giúp người mới vào dự án hiểu cấu trúc code, flow hoạt động, và biết phải tUser click ▶ Bắt Đầu
    └─ btnStart_Click() → ProcessImages() [async]   ← vòng lặp qua ảnh đã chọn, theo thứ tự
              ├─ CallGoogleVisionOCR()  ← gửi ảnh lên Google Vision (MainForm.cs)
              ├─ CleanOCRText()         ← lọc garbage lines (MainForm.cs)
              ├─ _ocrParsingService.ExtractAllFields()   ← parse 10 fields + Gemini fallback
              ├─ OCRInvoiceMapper.GetShipFee()           ← auto-fill TIỀN SHIP (3-tier lookup)
              ├─ OCRInvoiceMapper.GetNguoiDi()           ← auto-fill NGƯỜI ĐI (3-tier lookup)
              ├─ inject NGƯỜI ĐI / NGƯỜI LẤY từ UI (override nếu có)
              ├─ nếu thiếu field: fields["MISSING_FIELDS"] = "SHOP,MÃ,..." (tô đỏ khi xuất Excel)
              ├─ → append vào mappedDataList (INLINE — giữ đúng thứ tự quét, không có successList/failList)
              └─ → ghi raw OCR vào txtRawOCRLog, kết quả map vào txtProcessLog

User click 📊 Xuất Excel
    └─ ExportMappedDataToExcel()        ← user chọn file Excel, ghi vào sheet dd-MM
         └─ WriteDataRow()              ← tô đỏ từng cell có field trong MISSING_FIELDS ở đâu khi cần.

---

## 1. Tech Stack

| Layer | Technology |
|---|---|
| Runtime | .NET 8.0 Windows, C# |
| UI | Windows Forms (WinForms) |
| Excel I/O | [ClosedXML 0.102.3](https://github.com/ClosedXML/ClosedXML) |
| OCR | [Google Cloud Vision V1 3.8.0](https://cloud.google.com/vision) |
| AI Fallback | [Gemini API](https://aistudio.google.com/apikey) (free tier, Vision) — `Mscc.GenerativeAI` |
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
│   │   ├── OCRTextParsingService.cs   # Parse raw OCR text → extract 12 fields + Gemini fallback
│   │   ├── GeminiService.cs           # Gemini Vision fallback — đọc ảnh khi OCR parsing thiếu field
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
| `ExtractAllFields(text, out fields, geminiLog?)` | Public entry point — extract 10 fields, sau đó trigger Gemini fallback nếu còn thiếu field quan trọng |
| `ExtractAddressLine(text)` | Private — lấy dòng "địa chỉ:" **cuối cùng** hợp lệ (bỏ qua địa chỉ shop CN1/CN2) |
| `ExtractDistrictFromRawText(text)` | Private — fallback scan toàn bộ raw OCR tìm "Quận X" qua regex đa dòng; xử lý OCR wrap dòng giữa tên quận |
| `ExtractAmountLine(text, keywords)` | Private — tìm số tiền theo từ khoá; xử lý cả số cùng dòng lẫn số ở dòng tiếp theo |
| `NormalizeToThousands(raw)` | Private — chuẩn hóa về nghìn đồng (1,500,000 → 1500) |
| `ExtractDate(text)` | Private — parse ngày từ text |
| `RemoveDiacritics(s)` | Private static — bỏ dấu tiếng Việt, dùng bởi `ExtractDistrictFromRawText` |

**Gemini Fallback pipeline:**
```
OCR text parsing (regex)
    → nếu thiếu QUẬN: ExtractDistrictFromRawText() [không tốn quota]
    → nếu vẫn thiếu QUẬN / TÊN KH / MÃ / TIỀN THU / NGÀY LẤY:
         GeminiService.ParseInvoiceFromImageAsync() [đọc ảnh gốc]
              → thử tuần tự: flash-lite → 2.0-flash-lite → 2.0-flash → flash → pro
              → hết quota model nào → tự động sang model tiếp theo
```

**Edge cases đã xử lý (từ data thật):**

| Input thực tế | Vấn đề | Cách xử lý |
|---|---|---|
| `Địa Chi: 132 bên Vân đồn,p6,q4 - -` | OCR drop dấu `ỉ` → `"chi"` thay vì `"chỉ"` | Match thêm `"địa chi"` (có dấu `ị`) + `"dia chi"` (không dấu) |
| Hóa đơn có 2 dòng `Địa Chi/Chỉ:` (shop CN1 + khách hàng) | Parse nhầm địa chỉ shop | Lấy dòng **cuối cùng** hợp lệ; bỏ qua nếu chứa `CN\d / HOTLINE / SĐT` |
| `132 bên Vân đồn,p6,q4 - -` | Trailing garbage `- -` | Strip `[\s\-]+$` sau khi extract |
| `A25 hotel ( phòng 706) 184 nguyễn trãi, phường phạm ngũ lão, q1` | Số nhà phức tạp (tên khách sạn + số phòng + số nhà) | `ExtractHouseAndStreet` dùng greedy regex lấy đến số cuối cùng |
| `So HD: HD130781` (không dấu) | OCR drop dấu `ố` → `"So"` | Regex `So\s*H[ĐD]` đã cover |
| Số tiền trên dòng riêng (`Tổng tiền hàng:\n1,500,000`) | Số không cùng dòng keyword | `ExtractAmountLine` check thêm `lines[i+1]` |
| `TIỀN SHIP` không có trên hóa đơn | Field trống → lỗi validation | Không còn required — auto-fill từ bảng phí theo phường/quận (3-tier) |
| `363-365-367, 363 Đ. Hùng Vương - Khải Nam Transpost – –` | Số nhà là dãy số có `-`, tên business rác sau ` - ` | Strip ` - <tên không phải địa chỉ>` ở cuối; `Đ.` không bị strip vì được loại trừ khỏi regex |
| `Địa chỉ: 11 In Dung Vương Phường An Đông TP HCM ạ` | `"Phường An Đông"` và `"TP HCM"` cuối địa chỉ | Strip `Phường <tên>` + `TP HCM` ở cuối trước khi pass vào AddressParser |
| `Địa chỉ: ..., Phường 22, Quận B\nh Thạnh -` | OCR wrap tên quận qua 2 dòng | `ExtractDistrictFromRawText`: ghép text → regex → `AddressParser.Parse("q. Bình Thạnh")` |
| `THU 7.280+SHIP` | Bước 0 regex bắt "7.280" → NormalizeToThousands → chia /1000 → 7 (sai) | Bước 0 dùng digit-strip trực tiếp, không gọi NormalizeToThousands; "7.280" → 7280 ✅ |

### `ExcelInvoiceService`
**Mục đích:** Ghi dữ liệu OCR vào file Excel của khách (20 cột cố định)  
**File Excel:** hardcoded `"CHÂU NGÂN- THÁNG 2.2026- ĐỐI SOÁT.xlsx"` ⚠️  

| Method | Mô tả |
|---|---|
| `WriteInvoiceData(dataList, filePath, sheetName)` | Public entry point — ghi danh sách đơn vào sheet, tạo sheet nếu chưa có |
| `WriteDataRow(worksheet, row, data, ...)` | Ghi 1 row; tô đỏ nhạt các cell nằm trong `data["MISSING_FIELDS"]`; tô đỏ đậm nếu MÃ rỗng |
| `InvoiceExists(ma)` | Kiểm tra mã đơn đã tồn tại trong sheet chưa |

**Logic highlight thiếu field (WriteDataRow):**
```csharp
// data["MISSING_FIELDS"] = "SHOP,MÃ,TIỀN THU" (do OcrTab.cs ghi vào)
var missingSet = new HashSet<string>(data["MISSING_FIELDS"].Split(','), OrdinalIgnoreCase);

// fieldToCol map: "SHOP"→2, "TÊN KH"→3, "MÃ"→4, "ĐỊA CHỈ"→5, "QUẬN"→6,
//                 "TIỀN THU"→7, "TIỀN SHIP"→8, "NGÀY LẤY"→12, "GHI CHÚ"→13
foreach (var fieldName in missingSet)
    if (fieldToCol.TryGetValue(fieldName, out int col))
        worksheet.Cell(row, col).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFD0D0"); // đỏ nhạt

// MÃ rỗng → đỏ đậm (riêng biệt, luôn apply)
if (string.IsNullOrEmpty(ma))
    worksheet.Cell(row, COL_MA).Style.Fill.BackgroundColor = XLColor.FromHtml("#FF9999");
```

**Thứ tự xuất Excel:** `mappedDataList` được append inline trong vòng quét ảnh (không có successList/failList split) → thứ tự Excel = thứ tự ảnh quét.

### `GeminiService`
**Mục đích:** Fallback parser — khi `OCRTextParsingService` vẫn còn field thiếu sau regex, gửi ảnh gốc lên Gemini Vision để extract.  
**API key:** Lấy miễn phí tại https://aistudio.google.com/apikey — điền vào `AppConstants.GEMINI_API_KEY`.  
**Model fallback** (tự động thử tuần tự khi quota hết, quota nhiều → ít):

| Thứ tự | Model | Ghi chú |
|---|---|---|
| 1 | `gemini-2.5-flash-lite` | Quota nhiều nhất, nhanh nhất |
| 2 | `gemini-2.0-flash-lite` | Deprecated, còn đến Jun 2026 |
| 3 | `gemini-2.0-flash` | Deprecated, còn đến Jun 2026 |
| 4 | `gemini-2.5-flash` | Cân bằng |
| 5 | `gemini-2.5-pro` | Xịn nhất, quota ít nhất — last resort |

Khi gặp lỗi **429 / TooManyRequests / RESOURCE_EXHAUSTED** → tự động thử model tiếp theo.  
Lỗi khác (mất mạng, sai key) → báo ngay, không tiếp tục.

| Method | Mô tả |
|---|---|
| `ParseInvoiceFromImageAsync(imagePath)` | Gọi Gemini Vision, loop qua MODEL_FALLBACK_LIST, trả `(GeminiInvoiceResult, error)` |
| `IsConfigured` | `true` khi API key đã điền |

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
**Mục đích hiện tại:** Chứa model `OCRInvoiceData` và các helper tra phí ship / người đi.  

| Method / Class | Mô tả |
|---|---|
| `OCRInvoiceData` | Model class chứa tất cả fields của 1 invoice |
| `GetShipFee(phuong, quan)` | Tra phí ship — 3-tier: Phường → SHIPPING_FEES_BY_WARD → Phường→Quận via WARD_TO_DISTRICT_MAP → SHIPPING_FEES_BY_QUAN → Quận trực tiếp |
| `GetNguoiDi(phuong, quan)` | Tra người đi — 3-tier tương tự, dùng AREA_TO_NGUOI_DI |
| `NormalizeKey(s)` | Strip dấu + lowercase + expand alias viết tắt qua `_abbrevMap` |

**3-tier lookup (GetShipFee / GetNguoiDi):**
```
Tier 3 (phường cụ thể):    SHIPPING_FEES_BY_WARD[NormalizeKey(phuong)]
    ↓ miss
Tier 2.5 (phường → quận): WARD_TO_DISTRICT_MAP[NormalizeKey(phuong)] → SHIPPING_FEES_BY_QUAN[quan]
    ↓ miss
Tier 2 (quận trực tiếp):  SHIPPING_FEES_BY_QUAN[NormalizeKey(quan)]
```

**Alias expand (_abbrevMap trong NormalizeKey):**
```
"bh thanh" / "b thanh" / "bthanh" → "binh thanh"
"t binh"   → "tan binh"    "t phu"  → "tan phu"
"g vap"    → "go vap"      "b tan"  → "binh tan"
"t duc"    → "thu duc"     "p nhuan"→ "phu nhuan"
...
```

**Q8 phường-level split (SHIPPING_FEES_BY_WARD):**
- P.1–4, 8–10 (+ tên mới 2025): 25k
- P.5–7, 11–16 (+ tên mới 2025): 30k

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
| Cập nhật bảng phí ship theo phường (Q8 split...) | `AppConstants.SHIPPING_FEES_BY_WARD` |
| Thêm phường mới vào map phường→quận | `AppConstants.WARD_TO_DISTRICT_MAP` |
| Thêm alias viết tắt địa chỉ | `OCRInvoiceMapper._abbrevMap` trong `NormalizeKey()` |
| Thêm shared UI control style | `utils/UIHelper.cs` |
| Thêm search bar cho RichTextBox | `UIHelper.CreateRichTextBoxSearchBar()` |
| Thêm shared helper (dùng nhiều tab) | `MainForm.cs` |
| Thay đổi logic tính toán Excel Viewer | `InvoiceTab.cs` — `CalculateAllRows()` |
| Thay đổi cách detect header row | `InvoiceTab.cs` — `DetectHeaderRow()` + `AppConstants.HEADER_ROW_KEYWORDS` |
| Thay đổi cách OCR gọi Google | `MainForm.cs` — `CallPythonOCR()` |
| Thêm loại ảnh được chấp nhận | `MainForm.cs` — `GetImageFiles()` |
| Đổi model Gemini / thứ tự fallback | `GeminiService.MODEL_FALLBACK_LIST` |
| Đổi Gemini API key | `AppConstants.GEMINI_API_KEY` |

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
| 11 | `AppConstants.GEMINI_API_KEY` | API key Gemini nhúng thẳng | Không nên commit lên git public |

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
| `ExcelInvoiceService.cs` | constructor | Tên file Excel hardcoded theo tháng — cần đổi mỗi tháng |

---

## 10. Warnings hiện tại (không block build)

| Warning | Nguồn | Giải thích |
|---|---|---|
| `CS8669` (×6) | `MainForm.Designer.cs` | Nullable annotation trong auto-generated code — bỏ qua |
| `CS0618` | `MainForm.cs:57` | `GoogleCredential.FromFile()` deprecated — vẫn hoạt động, fix sau |

---

## 11. Commands hữu ích
### a. Build file .exe
``` 
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o publish\
```

### b. Build và run project
``` 
dotnet build
dotnet run
```

### c. Script rename images để dễ track
```
powershell -ExecutionPolicy Bypass -File ".\rename-images.ps1" -FolderPath "data\27-02-2026\images" -AutoConfirm
```
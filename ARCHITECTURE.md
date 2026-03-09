# TextInputter — Architecture Guide

> **Mục đích:** Tài liệu này giúp người mới vào dự án hiểu cấu trúc code, flow hoạt động, và biết phải tìm/sửa ở đâu khi cần.

---

## 1. Tech Stack

| Layer | Technology |
|---|---|
| Runtime | .NET 8.0 Windows, C# |
| UI | Windows Forms (WinForms) |
| Excel I/O | [ClosedXML 0.102.3](https://github.com/ClosedXML/ClosedXML) |
| OCR | [Google Cloud Vision V1 3.8.0](https://cloud.google.com/vision) |
| AI Fallback | [Gemini API](https://aistudio.google.com/apikey) (free tier, Vision) — `Mscc.GenerativeAI` |
| Credentials | `text-extractor-489011-ee19271357bd.json` (Google Service Account) |

---

## 2. Cấu trúc thư mục

```
TextInputter/
├── TextInputter.csproj
├── ARCHITECTURE.md
├── README.md
│
├── main/
│   ├── Program.cs               # Entry point
│   ├── AppConstants.cs          # TẤT CẢ hardcoded values tập trung tại đây
│   ├── MainForm.cs              # Shared fields + constructor + shared helpers
│   ├── MainForm.Designer.cs     # Form-level skeleton
│   ├── tabs/
│   │   ├── OcrTab.cs            # OCR tab: logic + InitializeOCRTab()
│   │   ├── InvoiceTab.cs        # Invoice tab: logic handlers
│   │   ├── InvoiceTab.UI.cs     # Invoice tab: control declarations
│   │   ├── ManualInputTab.cs    # Manual Input tab
│   │   └── UpdateTab.UI.cs      # Update tab UI (stub)
│   ├── Services/
│   │   ├── OCRTextParsingService.cs   # Parse raw OCR text → 12 fields + Gemini fallback
│   │   ├── GeminiService.cs           # Gemini Vision fallback (5 model auto-fallback)
│   │   ├── ExcelInvoiceService.cs     # Ghi/cập nhật invoice vào Excel (19 cột)
│   │   ├── OCRInvoiceMapper.cs        # Model + tra ship fee / người đi
│   │   └── AddressParser.cs           # Parse địa chỉ VN
│   └── utils/
│       ├── UIHelper.cs
│       └── AddressParsingDialog.cs
├── resources/app.ico
└── data/sample/{excel,images}
```

---

## 3. Pattern: Partial Classes (quan trọng!)

`MainForm` được **split thành nhiều file** bằng cơ chế `partial class` của C#.

| File | Chứa gì |
|---|---|
| `MainForm.Designer.cs` | Form-level: panelTop, panelLeft, panelBottom, tabMainControl, TabPages |
| `<Tab>.UI.cs` | Control field declarations + `Initialize<Tab>UI()` — chỉ layout, không event logic |
| `<Tab>.cs` | Event handlers, business logic, service calls |
| `MainForm.cs` | Shared fields, constructor, shared helpers |

Tất cả đều **cùng 1 class** — mọi field/method ở file nào cũng truy cập được từ file khác.

---

## 4. Shared Fields (MainForm.cs)

| Field | Type | Mô tả |
|---|---|---|
| `folderPath` | `string` | Đường dẫn folder ảnh OCR đang chọn |
| `imageFiles` | `List<string>` | Danh sách file ảnh trong folder |
| `isProcessing` | `bool` | Flag chống double-click khi đang xử lý |
| `visionClient` | `ImageAnnotatorClient` | Google Vision client |
| `_ocrParsingService` | `OCRTextParsingService` | Service parse OCR text |
| `mappedDataList` | `List<Dictionary<string,string>>` | Cache kết quả OCR đã map (để export) |

---

## 5. Flow chính

### 5A. OCR Batch (OcrTab.cs) — flow chính của dự án

```
User click 📁 Chọn Thư Mục
    └─ SelectOCRFolder() → GetImageFiles()   ← lọc .jpg/.png/.webp

User chọn Người Đi / Người Lấy (ComboBox)
User (tuỳ chọn) tick "Manual Người Đi/Lấy" để override auto-fill

User click ▶ Bắt Đầu
    └─ btnStart_Click() → ProcessImages() [async]
              ├─ CallGoogleVisionOCR()                ← Google Vision API
              ├─ CleanOCRText()                       ← lọc garbage lines
              ├─ _ocrParsingService.ExtractAllFields()  ← parse 12 fields
              ├─ AddressParser.Parse()                ← tách SoNha/TenDuong/Phuong/Quan
              ├─ OCRInvoiceMapper.GetShipFee(phuong, quan, duong)  ← auto-fill TIỀN SHIP (4-tier)
              ├─ OCRInvoiceMapper.GetNguoiDi(phuong, quan)         ← auto-fill NGƯỜI ĐI
              │     (nếu GetNguoiDi = "An Tam" → append " dd-MM" ngày hôm nay)
              ├─ inject NGƯỜI ĐI / NGƯỜI LẤY từ ComboBox (nếu manual mode)
              ├─ nếu thiếu field → fields["MISSING_FIELDS"] = "SHOP,MÃ,..."
              ├─ → append vào mappedDataList (giữ thứ tự quét)
              └─ → ghi raw OCR vào txtRawOCRLog, kết quả map vào txtProcessLog

User click 📊 Xuất Excel
    └─ ExportMappedDataToExcel()
         ├─ OpenFileDialog → user chọn file Excel đích
         └─ ExcelInvoiceService.ExportBatch()   ← ghi data rows vào sheet dd-MM

User click 🧮 Tính Tiền (sau khi đã xuất)
    └─ ExcelInvoiceService.ApplyFormulasAndSummary()
         ├─ Tính TIỀN HÀNG = TIỀN THU - TIỀN SHIP cho từng row
         ├─ Thêm bảng tổng trái (per SHOP): SUMIFS TIỀN THU, COUNTIFS đơn, -SUMIFS TIỀN SHIP
         └─ Thêm bảng tổng phải (per NGƯỜI ĐI): SUMIFS + COUNTIFS
```

### 5B. Excel Viewer + Daily Report (InvoiceTab.cs)

```
User click 📁 Mở File
    └─ BtnOpenExcel_Click() → LoadExcelFile()
         └─ DetectHeaderRow()     ← tìm header dựa vào HEADER_ROW_KEYWORDS
              └─ MapColumnIndices() ← gán cột SHOP, TIỀN THU, TIỀN SHIP...

User click 🧮 Tính
    └─ CalculateAllRows() → DisplayDailyReport()

User click 💾 Lưu
    └─ SaveDailyReportToExcel()   ← ghi DailyTotalReport.xlsx
```

### 5C. Manual Input (ManualInputTab.cs)

```
User điền fields vào form → SaveManualEntry()
    ⚠️ TODO: hiện chỉ hiện MessageBox — chưa ghi vào Excel
```

---

## 6. Services

### `OCRTextParsingService`
**Input:** raw OCR text (string từ Google Vision)  
**Output:** `Dictionary<string, string>` chứa các fields

| Method | Mô tả |
|---|---|
| `ExtractAllFields(text, out fields, geminiLog?)` | Public entry point — extract 12 fields, trigger Gemini fallback nếu thiếu |
| `ExtractAddressLine(text)` | Lấy dòng "địa chỉ:" **cuối cùng** hợp lệ (bỏ qua địa chỉ shop CN1/CN2) |
| `ExtractDistrictFromRawText(text)` | Fallback scan toàn bộ raw OCR tìm "Quận X"; xử lý OCR wrap dòng |
| `ExtractAmountLine(text, keywords)` | Tìm số tiền theo từ khoá; xử lý số ở dòng tiếp theo |
| `NormalizeToThousands(raw)` | Chuẩn hóa về nghìn đồng (1,500,000 → 1500) |

**Gemini Fallback pipeline:**
```
OCR text parsing (regex)
    → nếu thiếu QUẬN: ExtractDistrictFromRawText() [không tốn quota]
    → nếu vẫn thiếu QUẬN / TÊN KH / MÃ / TIỀN THU / NGÀY LẤY:
         GeminiService.ParseInvoiceFromImageAsync() [đọc ảnh gốc]
              → thử tuần tự: 2.5-flash-lite → 2.0-flash-lite → 2.0-flash → 2.5-flash → 2.5-pro
              → hết quota model nào → tự động sang model tiếp theo
```

**Edge cases đã xử lý (từ data thật):**

| Input thực tế | Vấn đề | Cách xử lý |
|---|---|---|
| `Địa Chi: 132 bên Vân đồn,p6,q4 - -` | OCR drop dấu `ị` → `"chi"` | Match thêm `"địa chi"` + `"dia chi"` |
| Hóa đơn có 2 dòng `Địa Chỉ:` (shop CN1 + khách) | Parse nhầm địa chỉ shop | Lấy dòng **cuối cùng** hợp lệ; bỏ qua nếu chứa `CN\d / HOTLINE / SĐT` |
| `So HD: HD130781` (không dấu) | OCR drop dấu `ố` | Regex `So\s*H[ĐD]` đã cover |
| Số tiền trên dòng riêng (`Tổng tiền hàng:
1,500,000`) | Số không cùng dòng keyword | `ExtractAmountLine` check thêm `lines[i+1]` |
| `Địa chỉ: ..., Phường 22, Quận Bình Thạnh -` | OCR wrap tên quận qua 2 dòng | `ExtractDistrictFromRawText`: ghép text → regex → AddressParser |
| `THU 7.280+SHIP` | "7.280" → NormalizeToThousands → 7 (sai) | Bước 0 dùng digit-strip trực tiếp; "7.280" → 7280 ✅ |

### `ExcelInvoiceService`
**Mục đích:** Ghi / cập nhật dữ liệu invoice vào file Excel của khách (19 cột cố định).  
**File Excel:** user chọn qua `OpenFileDialog` khi export — **không hardcode tên file**.

**Cột (1-based):**

| Col | Tên | Col | Tên |
|-----|-----|-----|-----|
| 1 | TÌNH TRẠNG | 11 | NGƯỜI LẤY |
| 2 | SHOP | 12 | NGÀY LẤY |
| 3 | TÊN KH | 13 | GHI CHÚ |
| 4 | MÃ | 14 | ỨNG TIỀN |
| 5 | ĐỊA CHỈ | 15 | HÀNG TỒN |
| 6 | QUẬN | 16 | FAIL |
| 7 | TIỀN THU | 17 | COL1 |
| 8 | TIỀN SHIP | 18 | COL2 |
| 9 | TIỀN HÀNG | 19 | COL3 |
| 10 | NGƯỜI ĐI | | |

**Layout sheet:**
- Row 1: header cột (AutoFilter bật)
- Row 2: `THU x / NGAY x-x` (info ngày)
- Row 3+: data rows (DATA_START_ROW = 3)

**Methods chính:**

| Method | Mô tả |
|---|---|
| `ExportBatch(dataList, sheetName, sheetDate)` | Ghi batch data rows vào sheet, tạo sheet + header nếu chưa có. Trả `(added, updated)` |
| `ApplyFormulasAndSummary(sheetName, sheetDate)` | Tính TIỀN HÀNG, thêm bảng tổng trái (per SHOP) + bảng tổng phải (per NGƯỜI ĐI) |
| `InvoiceExists(soHoaDon, out existingSheet)` | Kiểm tra mã đơn đã tồn tại chưa (scan tất cả sheet) |
| `GetAllInvoiceNumbers()` | Lấy tất cả mã đơn đã có trong file |

**Logic TÌNH TRẠNG:**
- Đơn có MÃ → giữ nguyên TÌNH TRẠNG từ OCR (vd: "Chuyển Tiền", "COD")
- Đơn không có MÃ (hàng sỉ, ship gộp) → ghi `"hàng sỉ"`; không tô đỏ MÃ
- MÃ rỗng + không phải hàng sỉ → tô đỏ đậm cell MÃ

**Logic TIỀN THU:**
- TIỀN THU ghi vào Excel = `Tổng thanh toán (từ OCR) + TIỀN SHIP`
- TIỀN HÀNG (sau ApplyFormulas) = `TIỀN THU - TIỀN SHIP`

**Logic bảng tổng (ApplyFormulasAndSummary):**
```
Bảng trái (per SHOP):
  Row 1: Tên shop | SUMIFS(TIỀN THU, cột SHOP, "<tên shop>") | COUNTIFS(cột SHOP, "<tên shop>")
  Row 2:           | -SUMIFS(TIỀN SHIP, cột SHOP, "<tên shop>")     |

Bảng phải (per NGƯỜI ĐI):
  Tương tự, group theo cột NGƯỜI ĐI
```

**Logic tô màu:**
- data["MISSING_FIELDS"] = "SHOP,MÃ,..." → tô đỏ nhạt (#FFD0D0) các cell tương ứng
- MÃ rỗng + không phải hàng sỉ → tô đỏ đậm (#FF9999)

### `GeminiService`
**Mục đích:** Fallback parser — gửi ảnh gốc lên Gemini Vision khi regex vẫn thiếu field.  
**API key:** Điền vào `AppConstants.GEMINI_API_KEY`.

**Model fallback tự động (quota nhiều → ít):**

| Thứ tự | Model | Ghi chú |
|--------|-------|---------|
| 1 | `gemini-2.5-flash-lite` | Quota nhiều nhất, nhanh nhất |
| 2 | `gemini-2.0-flash-lite` | Deprecated, còn đến Jun 2026 |
| 3 | `gemini-2.0-flash` | Deprecated, còn đến Jun 2026 |
| 4 | `gemini-2.5-flash` | Cân bằng |
| 5 | `gemini-2.5-pro` | Xịn nhất, quota ít nhất — last resort |

Gặp lỗi **429 / RESOURCE_EXHAUSTED** → tự động thử model tiếp theo.

### `AddressParser`
**Input:** string địa chỉ thô  
**Output:** `ParsedAddress { SoNha, TenDuong, Phuong, Quan, Confidence }`

**Edge cases đã xử lý:**

| Input thực tế | Vấn đề | Cách xử lý |
|---|---|---|
| `5/1 phùng văn cung p2 phủ nhuận` | Không có dấu phẩy | Tự chèn phẩy trước `p<số>`, `q<số>` |
| `11 In Dung Vương` | Số nhà `11` bị nhận nhầm là Q.11 | Bare number chỉ match quận khi **toàn segment là số đó** |
| `363 Đ. Hùng Vương` | `Đ.` viết tắt Đường | Regex riêng bắt `<số> Đ. <tên>` |
| `phủ nhuận` / `phú nhuật` (OCR sai dấu) | Không match exact | Fuzzy: xóa dấu → match `"phu nhuan"` |
| `Tân Phú, TP Thủ Đức` | "TP Thủ Đức" bị parse nhầm | `DistrictAliasDict`: `"tp thu duc"` → `"thu duc"` |

### `OCRInvoiceMapper`

| Method | Mô tả |
|---|---|
| `GetShipFee(phuong, quan, duong?)` | Tra phí ship — 4-tier lookup (xem bên dưới) |
| `GetShipFeeByQuan(quan)` | Shortcut: chỉ tra theo quận |
| `GetNguoiDi(phuong, quan)` | Tra người đi — 3-tier tương tự, dùng `AREA_TO_NGUOI_DI` |
| `NormalizeKey(s)` | Strip dấu + lowercase + expand alias viết tắt qua `_abbrevMap` |

**4-tier ship fee lookup (`GetShipFee`):**
```
Tier 3   (phường):       SHIPPING_FEES_BY_WARD[NormalizeKey(phuong)]
    ↓ miss
Tier 2.8 (tên đường):    SHIPPING_FEES_BY_STREET[NormalizeKey(duong)]   ← partial match
    ↓ miss
Tier 2.5 (phường→quận):  WARD_TO_DISTRICT_MAP[phuong] → SHIPPING_FEES_BY_QUAN[mappedQuan]
    ↓ miss
Tier 2   (quận):         SHIPPING_FEES_BY_QUAN[NormalizeKey(quan)]
```

**Bảng ship fee hiện tại (`SHIPPING_FEES_BY_QUAN`):**

| Quận | Ship | Ghi chú |
|------|------|---------|
| Q1, Q3 | 20k | Gần trung tâm |
| Q4, Q5, Q10, Q11 | 25k | |
| Q6 | 25k base | đường Đặng Nguyên Cẩn → 30k (SHIPPING_FEES_BY_STREET) |
| Q7, Q2, Q12 | 30k | |
| Q8 | 25k base | P.5–7, P.11–16 → 30k (SHIPPING_FEES_BY_WARD) |
| Q9 | 30k base | Vinhome Grand Park → 35k (SHIPPING_FEES_BY_STREET) |
| Bình Thạnh, Phú Nhuận | 20k | |
| Gò Vấp, Tân Phú | 25k base | |
| Tân Bình | 25k base | đường Yên Thế, Quách Văn Tuấn → 30k (SHIPPING_FEES_BY_STREET) |
| Bình Tân, Thủ Đức | 30k | |
| Bình Chánh, Hóc Môn, Nhà Bè | 35k | |
| Củ Chi | 40k | |
| Cần Giờ | 50k | |

**Phân công người đi (`AREA_TO_NGUOI_DI`):**

| Quận/Khu | Người đi |
|----------|---------|
| Q1, Q3, Q4, Q5, Q10, Phú Nhuận, Tân Bình | c.hieu |
| Bình Thạnh, Thủ Đức, Gò Vấp, Q2 | c.cuong |
| Q9 | a.quyen |
| Còn lại (Q6, Q7, Q8, Q11, Q12, Tân Phú, Bình Tân...) | An Tam + ngày `dd-MM` |

**Alias expand (`_abbrevMap` trong `NormalizeKey`):**
```
"bh thanh" / "bthanh"  → "binh thanh"
"t binh"               → "tan binh"
"g vap"                → "go vap"
"t duc"                → "thu duc"
"p nhuan"              → "phu nhuan"
"t phu"                → "tan phu"
"b tan"                → "binh tan"
```

---

## 7. Thêm tính năng mới — làm ở đâu?

| Muốn làm gì | File cần edit |
|---|---|
| Thêm/sửa phí ship theo quận | `AppConstants.SHIPPING_FEES_BY_QUAN` |
| Thêm phí ship override theo phường | `AppConstants.SHIPPING_FEES_BY_WARD` |
| Thêm phí ship override theo tên đường cụ thể | `AppConstants.SHIPPING_FEES_BY_STREET` |
| Thêm phường mới vào map phường→quận | `AppConstants.WARD_TO_DISTRICT_MAP` |
| Thay đổi phân công người đi theo khu vực | `AppConstants.AREA_TO_NGUOI_DI` |
| Thêm alias viết tắt địa chỉ | `OCRInvoiceMapper._abbrevMap` trong `NormalizeKey()` |
| Thêm field mới vào OCR output | `OCRTextParsingService.ExtractAllFields()` |
| Thêm cột mới vào Excel export | `ExcelInvoiceService` (cập nhật COL_* constants + AddHeaderRow + WriteDataRow) |
| Thêm config/constant (data thuần) | `AppConstants.cs` |
| Thêm tab mới | Tạo `tabs/NewTab.cs` với `partial class MainForm` |
| Thay đổi logic tính toán Excel Viewer | `InvoiceTab.cs` — `CalculateAllRows()` |
| Đổi model Gemini / thứ tự fallback | `GeminiService.MODEL_FALLBACK_LIST` |
| Đổi Gemini API key | `AppConstants.GEMINI_API_KEY` |
| Thêm shared UI control style | `utils/UIHelper.cs` |
| Thêm shared helper (dùng nhiều tab) | `MainForm.cs` |

---

## 8. Danh sách Hardcoded cần cải thiện

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
| 9 | `AppConstants.GOOGLE_CREDENTIAL_FILE` | `"text-extractor-489011-ee19271357bd.json"` | Credential file cứng cạnh .exe |
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
| `AppConstants.AREA_TO_NGUOI_DI` | Tân Bình | Cả quận → c.hieu, nhưng thực tế chỉ một số phường — cần ward-level sau khi AddressParser phân biệt được |
| `SHIPPING_FEES_BY_STREET` | Gò Vấp xa | đường Quang Trung số lớn (401+) = 30k nhưng "quang trung" xuất hiện ở nhiều quận — cần logic quận+đường |

---

## 10. Warnings hiện tại (không block build)

| Warning | Nguồn | Giải thích |
|---|---|---|
| `CS8669` (×6) | `MainForm.Designer.cs` | Nullable annotation trong auto-generated code — bỏ qua |
| `CS0618` | `MainForm.cs` | `GoogleCredential.FromFile()` deprecated — vẫn hoạt động |

---

## 11. Commands hữu ích

### Build & run
```powershell
dotnet build
dotnet run
```

### Build file .exe standalone
```powershell
dotnet publish -c Release -r win-x64 --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -o publish\
```

### Rename ảnh để dễ track
```powershell
powershell -ExecutionPolicy Bypass -File ".\rename-images.ps1" `
  -FolderPath "data\Mar\images" -AutoConfirm
```

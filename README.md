# TextInputter

Ứng dụng Windows (C# WinForms) hỗ trợ nhân viên giao hàng nhập liệu hóa đơn nhanh từ ảnh chụp, xuất thẳng vào file Excel theo định dạng của khách.

---

## ⚠️ BƯỚC QUAN TRỌNG: Setup Google Cloud Credentials

Chương trình cần **Google Cloud service account credentials** để hoạt động.
### 1️⃣ Tạo Google Cloud Project

1. Truy cập: https://console.cloud.google.com
2. Tạo project mới (Project Name: `TextInputter` hoặc tùy ý)
3. Bật **Vision API**:
   - Menu → APIs & Services → Library
   - Search: "Cloud Vision API"
   - Click → Enable
4. Bật **Billing** (Google cung cấp 1000 requests/tháng miễn phí):
   - Menu → Billing
   - Link tài khoản billing

### 2️⃣ Tạo Service Account Credentials

1. Vào: APIs & Services → Credentials
2. Click: Create Credentials → Service Account
3. Điền thông tin:
   - Service account name: `textinputter-ocr`
   - Click: Create and Continue
4. Tạo Key:
   - Service Account → Keys tab
   - Add Key → Create new key
   - Format: **JSON**
   - Download file JSON (ví dụ: `text-extractor-489011-ee19271357bd.json`)

### 3️⃣ Copy vào project

- Đặt file JSON vào **gốc project**:
  ```
  d:\Work\Freelance\TextInputter\[tên-file-credentials].json
  ```

- **HOẶC** rename thành tên mặc định:
  ```
  text-extractor-489011-ee19271357bd.json
  ```

### 4️⃣ ⚠️ Thêm vào .gitignore (ĐẬU BẮT BUỘC!)

File credentials chứa **private key** → **KHÔNG được public lên GitHub**

Kiểm tra `.gitignore` có dòng này không:
```gitignore
text-extractor-489011-ee19271357bd.json
```

Nếu chưa có, thêm vào `.gitignore`

---

## 🤖 (Tuỳ chọn) Setup Gemini AI Fallback

Khi OCR parsing không đủ field (địa chỉ bị wrap dòng, quận không rõ...), app tự gửi ảnh lên **Gemini Vision** để đọc lại.

### Lấy API key miễn phí:
1. Truy cập: https://aistudio.google.com/apikey
2. Tạo API key mới (không cần billing)
3. Mở `main/AppConstants.cs`, điền key vào:
   ```csharp
   public const string GEMINI_API_KEY = "YOUR_KEY_HERE";
   ```

### Model fallback tự động (quota nhiều → ít):
```
gemini-2.5-flash-lite → gemini-2.0-flash-lite → gemini-2.0-flash → gemini-2.5-flash → gemini-2.5-pro
```
Hết quota model nào → tự động thử model tiếp theo.

> ⚠️ Để trống `""` = tắt Gemini, chỉ dùng rule-based parser.  
> ⚠️ Không commit API key lên git nếu repo public.

---

## 📝 File Sample Credentials

Sử dụng template trong `textinputter-google-credential-sample.json` để guide người khác setup:

```json
{
  "type": "service_account",
  "project_id": "textinputter",
  "private_key_id": "{private_key_id}",
  "private_key": "-----BEGIN PRIVATE KEY-----\n{private_key}\n-----END PRIVATE KEY-----\n",
  "client_email": "textinputter-ocr@textinputter.iam.gserviceaccount.com",
  "client_id": "{client_id}",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/...",
  "universe_domain": "googleapis.com"
}
```

**Thay đổi các trường:**
- `{private_key_id}` → Lấy từ file JSON download
- `{private_key}` → Lấy từ file JSON download (toàn bộ private key)
- `{client_id}` → Lấy từ file JSON download

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
d:\Work\Freelance\TextInputter\
├── main/
│   ├── AppConstants.cs          # Config tập trung: API keys, bảng phí ship, màu sắc...
│   ├── MainForm.cs              # Shared fields + constructor
│   ├── MainForm.Designer.cs     # Form skeleton
│   ├── Program.cs               # Entry point
│   ├── tabs/
│   │   ├── OcrTab.cs            # OCR batch tab
│   │   ├── InvoiceTab.cs        # Excel viewer + Daily Report
│   │   ├── InvoiceTab.UI.cs     # Invoice UI controls
│   │   └── ManualInputTab.cs    # Manual input tab
│   ├── Services/
│   │   ├── OCRTextParsingService.cs  # Parse OCR text → 12 fields + Gemini fallback
│   │   ├── GeminiService.cs          # Gemini Vision AI (5 model fallback)
│   │   ├── AddressParser.cs          # Parse địa chỉ VN
│   │   ├── ExcelInvoiceService.cs    # Ghi Excel
│   │   └── OCRInvoiceMapper.cs       # Model + ship fee lookup
│   └── utils/
│       ├── UIHelper.cs               # WinForms factory + search
│       └── AddressParsingDialog.cs   # Dialog xác nhận địa chỉ
├── resources/
│   └── app.ico
├── data/sample/                 # File mẫu để test
├── ARCHITECTURE.md              # Chi tiết kiến trúc, flow, edge cases
├── TextInputter.csproj          # Project file
├── text-extractor-489011-ee19271357bd.json              # ⚠️ Credentials Google (KHÔNG push)
└── textinputter-google-credential-sample.json  # Template sample
```

> Xem `ARCHITECTURE.md` để biết chi tiết flow, services, edge cases và hướng dẫn thêm tính năng.

---

## 📄 License

Miễn phí sử dụng - TextInputter OCR

---

## Hành vi đặc biệt

- **Hàng sỉ / ship gộp** (không có MÃ): TÌNH TRẠNG = `"hàng sỉ"`, không tô đỏ cột MÃ
- **Người đi "An Tam"**: tự động append ngày `dd-MM` vào cuối (vd: `"An Tam 15-03"`)
- **Duplicate MÃ**: nếu mã đơn đã tồn tại trong bất kỳ sheet nào → **cập nhật** thay vì thêm mới
- **TIỀN THU** trong Excel = Tổng thanh toán OCR + TIỀN SHIP (để tính đúng doanh thu)
- **TIỀN HÀNG** = TIỀN THU - TIỀN SHIP (tính sau khi bấm "Tính Tiền")


## Workflow/Business
                   ┌───────────────────────────────┐
                   │  BƯỚC 1: NHẬP HÀNG LẤY        │
                   │  (nhập tất cả đơn từ OCR/ảnh) │
                   └──────────┬────────────────────┘
                              │
                   ┌──────────▼───────────────────┐
                   │  BƯỚC 2: NHẬP NGƯỜI ĐI       │
                   │  (gán shipper cho từng đơn)  │
                   └──────────┬───────────────────┘
                              │
           ┌──────────────────▼──────────────────────┐
           │  BƯỚC 3: NHẬP HÀNG TRẢ (2 loại)         │
           │                                         │
           │  Loại 1 — Trả TRONG NGÀY:               │
           │    → ỨNG TIỀN = x, FAIL = xx            │
           │    (VD: row 3, trinh, 840k)             │
           │                                         │
           │  Loại 2 — Trả NGÀY TRƯỚC (hàng tồn):    │
           │    → Tìm trong bảng đối soát            │
           │    → Copy vào sheet hiện tại            │
           │    → ỨNG TIỀN = x, HÀNG TỒN = x,        │
           │      FAIL = xx                          │
           │    (VD: row 30, duyen my, 850k)         │
           └──────────────────┬──────────────────────┘
                              │
                   ┌──────────▼───────────────────┐
                   │  BƯỚC 4: TÍNH TIỀN           │
                   └──────────────────────────────┘
# 📄 TextInputter - Ứng dụng OCR hóa đơn tiếng Việt

Ứng dụng **Windows WinForms** để quét, nhận diện và trích xuất thông tin từ hình ảnh hóa đơn tiếng Việt với độ chính xác cực kỳ cao (99%+) nhờ **Google Cloud Vision API**, kết hợp **Gemini Vision AI** làm fallback khi parse địa chỉ thất bại.

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
   - Download file JSON (ví dụ: `textinputter-4a7bda4ef67a.json`)

### 3️⃣ Copy vào project

- Đặt file JSON vào **gốc project**:
  ```
  d:\Work\Freelance\TextInputter\[tên-file-credentials].json
  ```

- **HOẶC** rename thành tên mặc định:
  ```
  textinputter-4a7bda4ef67a.json
  ```

### 4️⃣ ⚠️ Thêm vào .gitignore (ĐẬU BẮT BUỘC!)

File credentials chứa **private key** → **KHÔNG được public lên GitHub**

Kiểm tra `.gitignore` có dòng này không:
```gitignore
textinputter-4a7bda4ef67a.json
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

## 🚀 Chạy chương trình

### Yêu cầu:
- ✅ .NET 8.0 SDK
- ✅ File credentials JSON đã copy vào folder
- ✅ Google Cloud Vision API đã enable
- ✅ Billing đã setup

### Chạy:
```powershell
cd d:\Work\Freelance\TextInputter
dotnet run
```

### Quy trình sử dụng:
1. **OCR Tab:** Chọn folder ảnh hóa đơn → nhập Người Đi / Người Lấy → Bắt Đầu
2. App gửi từng ảnh lên Google Vision → extract text → parse 12 fields (SHOP, TÊN KH, MÃ, địa chỉ, tiền, ngày...)
3. Nếu thiếu field: tự động fallback Gemini Vision đọc ảnh gốc. Nếu vẫn thiếu → đơn vẫn được xuất, các cell thiếu tô đỏ để điền tay
4. Kết quả hiện ở log theo **đúng thứ tự ảnh đã quét** → Xuất Excel
5. **Invoice Tab:** Mở file Excel của khách → Tính → xem Daily Report → Lưu báo cáo

---

## 📦 Build standalone .exe (Optional)

```powershell
dotnet publish -c Release -r win-x64 --self-contained true
```

Output `.exe`:
```
bin/Release/net8.0-windows/publish/TextInputter.exe
```

⚠️ **Lưu ý:** File credentials vẫn cần có trong cùng folder với `.exe`

---

## ✨ Tính năng:

✅ **OCR hàng loạt** — Batch process nhiều ảnh hóa đơn cùng lúc  
✅ **Nhận diện chính xác** — Google Vision API (99%+)  
✅ **Parse thông minh** — Tự động extract 12 fields: tên KH, mã HĐ, địa chỉ, tiền thu, tiền ship, ngày...  
✅ **Gemini AI Fallback** — Khi regex fail → gửi ảnh lên Gemini Vision, tự chuyển model khi hết quota  
✅ **Địa chỉ VN** — Tách SỐ NHÀ / TÊN ĐƯỜNG / PHƯỜNG / QUẬN, cover sáp nhập ĐVHC TP.HCM 2025  
✅ **Auto phí ship** — Tra bảng phí theo phường/quận (Q8: split từng phường; các quận khác: tra theo quận)  
✅ **Alias địa chỉ** — Nhận dạng viết tắt như "bh thanh" → "bình thạnh", "t binh" → "tân bình"...  
✅ **Thứ tự quét** — Excel xuất đúng thứ tự ảnh đã quét, không đảo lộn  
✅ **Highlight thiếu field** — Đơn thiếu field vẫn xuất, tô đỏ các cell cần điền tay (không còn row FAIL)  
✅ **Excel export** — Xuất ra sheet theo ngày, ghi đúng 20 cột template  
✅ **Daily Report** — Tổng hợp doanh thu, tiền ship, số đơn theo ngày  
✅ **UI tiếng Việt** — Search log, màu sắc trực quan

---

## 💰 Chi phí

**Google Cloud Vision API:**
- Miễn phí 1,000 requests/tháng
- Sau đó: $0.6 per 1,000 requests
- Ví dụ: 1,000 ảnh/tháng ≈ $0.6

**Gemini Vision AI:**
- Hoàn toàn **miễn phí** (free tier) với API key từ https://aistudio.google.com/apikey
- 5 model fallback tự động — chỉ dùng khi OCR parsing không đủ field

---

## 🛠️ Troubleshooting

### ❌ "PermissionDenied: This API method requires billing to be enabled"
**Nguyên nhân:** Billing chưa setup  
**Fix:** Vào Google Cloud Console → Billing → Link tài khoản

### ❌ "Could not find credentials"
**Nguyên nhân:** File JSON không ở đúng vị trí  
**Fix:** Kiểm tra file `.json` nằm trong folder project gốc

### ❌ "Vision API not enabled"
**Nguyên nhân:** API chưa được bật  
**Fix:** APIs & Services → Library → Cloud Vision API → Enable

### ❌ "Invalid JSON in credentials"
**Nguyên nhân:** File JSON bị lỗi  
**Fix:** Download file mới từ Google Cloud Console

### ❌ Gemini: "Quota exceeded" / "TooManyRequests"
**Nguyên nhân:** Hết free quota của model đang dùng  
**Fix:** App tự động fallback — không cần làm gì. Nếu tất cả 5 model đều hết → chờ reset quota (12:00 AM Pacific time) hoặc chạy lại ngày hôm sau.

---

## 📂 Cấu trúc Project

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
├── textinputter-4a7bda4ef67a.json              # ⚠️ Credentials Google (KHÔNG push)
└── textinputter-google-credential-sample.json  # Template sample
```

> Xem `ARCHITECTURE.md` để biết chi tiết flow, services, edge cases và hướng dẫn thêm tính năng.

---

## 📄 License

Miễn phí sử dụng - TextInputter OCR

---

## 💡 Ghi chú quan trọng

- **✅ Google credentials KHÔNG commit lên GitHub** — Đã thêm vào `.gitignore`
- **✅ Gemini API key KHÔNG commit** — Điền vào `AppConstants.cs` nhưng không push nếu repo public
- **✅ Sử dụng template `textinputter-google-credential-sample.json`** để guide người khác cách setup
- **✅ Mỗi service account credentials khác nhau** — Thay đổi theo Google Cloud project của mình


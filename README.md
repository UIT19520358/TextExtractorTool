# 📄 TextInputter - Ứng dụng OCR tiếng Việt với Google Cloud Vision API

Ứng dụng **Windows WinForms** để quét, nhận diện và trích xuất văn bản tiếng Việt từ hình ảnh với độ chính xác cực kỳ cao (99%+) nhờ **Google Cloud Vision API**.

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
1. Chọn folder chứa ảnh (hoặc drag-drop)
2. Chương trình quét tất cả ảnh: `.jpg`, `.png`, `.jpeg`, `.bmp`
3. Google Vision API nhận diện chữ từng ảnh
4. Hiển thị kết quả OCR lên UI
5. Có thể lưu kết quả hoặc in

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

✅ **Quét hàng loạt** - Process nhiều ảnh cùng lúc  
✅ **Nhận diện chính xác** - Google Vision API (99%+)  
✅ **Hỗ trợ tiếng Việt** - Chữ Việt, dấu thanh (á, à, ả, ã, ạ...)  
✅ **Lọc rác** - Tự động xóa text không hợp lệ  
✅ **UI thân thiện** - Vietnamese UI, nút màu sắc  
✅ **Lưu kết quả** - Export text to file  

---

## 💰 Chi phí

**Google Cloud Vision API pricing:**
- **1-1,000,000 requests/tháng**: $0.6 per 1,000 requests (miễn phí 1,000 requests/tháng)
- Ví dụ: 1,000 ảnh ≈ $0.6/tháng

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

---

## 📂 Cấu trúc Project

```
d:\Work\Freelance\TextInputter\
├── main/
│   ├── MainForm.cs              # UI chính
│   ├── MainForm.Designer.cs     # Design form
│   └── Program.cs               # Entry point
├── images/                       # Ảnh test
├── bin/                         # Build output
├── obj/                         # Build temp
├── .gitignore                   # Ignore credentials (quan trọng!)
├── .vscode/
│   └── tasks.json               # Build tasks
├── README.md                    # File này
├── TextInputter.csproj          # Project file
├── textinputter-4a7bda4ef67a.json              # ⚠️ Credentials (KHÔNG push)
└── textinputter-google-credential-sample.json  # Template sample
```

---

## 📄 License

Miễn phí sử dụng - TextInputter OCR

---

## 💡 Ghi chú quan trọng

- **✅ Credentials KHÔNG được commit lên GitHub** - Đã thêm vào `.gitignore`
- **✅ Sử dụng template `textinputter-google-credential-sample.json`** để guide người khác cách setup
- **✅ Mỗi service account credentials khác nhau** - Thay đổi theo Google Cloud project của mình


using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Mscc.GenerativeAI;

namespace TextInputter.Services
{
    /// <summary>
    /// Fallback parser dùng Gemini Vision khi AddressParser không tách được QUẬN.
    /// Gửi ảnh gốc lên Gemini → extract toàn bộ fields hóa đơn → điền vào chỗ còn thiếu.
    ///
    /// Chỉ được gọi khi confidence thấp để tiết kiệm quota free (1500 req/ngày).
    /// API key lấy miễn phí tại: https://aistudio.google.com/apikey
    /// </summary>
    public class GeminiService
    {
        private readonly string _apiKey;

        // Thứ tự ưu tiên: quota nhiều nhất → ít nhất, tự fallback khi quota hết.
        // Tất cả đều hỗ trợ Vision (đọc ảnh), có free tier, shutdown sớm nhất Jun 2026.
        // ─────────────────────────────────────────────────────────────────────────────
        // gemini-2.5-flash-lite : quota nhiều nhất, nhanh nhất  → dùng trước
        // gemini-2.0-flash-lite : deprecated nhưng còn đến Jun 2026
        // gemini-2.0-flash      : deprecated nhưng còn đến Jun 2026
        // gemini-2.5-flash      : cân bằng tốt
        // gemini-2.5-pro        : xịn nhất, quota ít nhất       → dùng sau cùng
        // ─────────────────────────────────────────────────────────────────────────────
        private static readonly string[] MODEL_FALLBACK_LIST = new[]
        {
            "gemini-2.5-flash-lite", // ~nhiều nhất — fast, high-volume
            "gemini-2.0-flash-lite", // deprecated fallback
            "gemini-2.0-flash", // deprecated fallback
            "gemini-2.5-flash", // cân bằng
            "gemini-2.5-pro", // xịn nhất, quota ít nhất — last resort
        };

        /// <summary>
        /// Toàn bộ fields extract được từ 1 ảnh hóa đơn qua Gemini Vision.
        /// Chỉ điền field nào Gemini đọc được; field nào không rõ để chuỗi rỗng.
        /// </summary>
        public class GeminiInvoiceResult
        {
            // Địa chỉ
            public string DiaChi { get; set; } = "";
            public string Phuong { get; set; } = "";
            public string Quan { get; set; } = "";

            // Invoice fields
            public string TenShop { get; set; } = ""; // Tên shop/cửa hàng (header hóa đơn)
            public string TenKH { get; set; } = "";
            public string Ma { get; set; } = ""; // Số HĐ, VD: HD0123
            public string TienThu { get; set; } = ""; // Tổng tiền thu (đã trừ giảm giá)
            public string TienShip { get; set; } = ""; // Phí vận chuyển
            public string NgayLay { get; set; } = ""; // Ngày lấy/giao hàng

            // Loại đơn: "COD" | "SHIP_ONLY_FREE" | "SHIP_ONLY_PAID"
            // Gemini detect từ keyword trên nhãn: "KHÔNG THU SHIP", "THU SHIP", "THU X+SHIP"
            public string InvoiceType { get; set; } = ""; // rỗng = Gemini không detect được, dùng OCR logic
        }

        // Prompt yêu cầu Gemini đọc ảnh hoá đơn/nhãn ship và trả về đầy đủ fields dạng JSON
        private const string PROMPT =
            @"Đây là ảnh hóa đơn hoặc nhãn vận chuyển (shipping label) tiếng Việt của shop thời trang tại TP.HCM.
Có thể là một trong các loại sau:
  [LOẠI A] Hóa đơn shop in sẵn: có tên shop lớn ở đầu (VD: ĐOÀN NGÂN CHÂU), có 'Khách hàng:', 'Địa chỉ:', 'Số HĐ:', số tiền cuối cùng là tiền thu thực tế.
  [LOẠI B] Nhãn ship có sticker địa chỉ in: có nhãn 'Địa chỉ mới:' hoặc địa chỉ nằm trong ô sticker riêng, tên khách thường viết tay bên ngoài sticker, HOTLINE là số của shop (không phải khách), không có số HĐ — để ma=''.
  [LOẠI C] Nhãn ship viết tay: tên khách + bí danh viết to ở trên (VD: 'Thương 1994'), địa chỉ giao nằm trong sticker hoặc viết bên dưới, số điện thoại kèm tên khách, không có số HĐ — để ma=''.
  [LOẠI D] Nhãn ship có 'Khách hàng:' nhưng không có số HĐ: địa chỉ gồm số nhà + đường + phường + quận viết liền, số điện thoại có label 'SỐ ĐIỆN THOẠI:', tiền thu có dạng 'THU X,XXX + SHIP' hoặc 'THU X,XXX+SHIP'.

Hãy đọc kỹ ẢNH GỐC và trả về JSON theo đúng format sau, không giải thích thêm:
{
  ""ten_shop"": ""<tên shop/cửa hàng, thường là dòng IN HOA lớn ở đầu hóa đơn. Với nhãn ship (Loại B/C/D) thường không rõ — để trống>"",
  ""ten_kh"": ""<tên khách hàng. Loại A: dòng 'Khách hàng:'. Loại B: tên viết tay trên/ngoài sticker hoặc tên trong sticker. Loại C: tên/biệt danh viết to ở trên (VD: 'Thương 1994'). Loại D: dòng 'Khách hàng:'. KHÔNG lấy số điện thoại làm tên>"",
  ""ma"": ""<số hóa đơn, thường dạng HD + số, VD: HD0123. Nhãn ship (Loại B/C/D) không có → để trống>"",
  ""dia_chi"": ""<toàn bộ địa chỉ GIAO HÀNG (địa chỉ KHÁCH HÀNG) trước phần phường/quận — gồm số nhà + tên đường, hoặc block/căn hộ + tên chung cư + tên đường nếu có. KHÔNG bao gồm phường/quận. VD: '68 Nguyễn Trãi' hoặc 'Block D2, Chung cư Sài Gòn Riverside, 4 Đào Trí' hoặc 'Tòa S501, Vinhome Grand Park' hoặc '401 Quang Trung'>"",
  ""phuong"": ""<phường/xã của địa chỉ GIAO HÀNG, nếu có. VD: Phường 3, Phường 22, An Hội Tây, Long Thạnh Mỹ, Phú Lâm>"",
  ""quan"": ""<quận/huyện của địa chỉ GIAO HÀNG — QUAN TRỌNG: đọc kỹ toàn bộ địa chỉ để tìm quận, kể cả khi viết tắt. Chỉ ghi TÊN QUẬN hoặc SỐ QUẬN thuần túy, KHÔNG ghi 'Quận' hay 'Q.' phía trước, KHÔNG ghi tên phường. VD: 1, 4, 10, Bình Thạnh, Tân Phú, Gò Vấp, Thủ Đức, Phú Nhuận>"",
  ""tien_thu"": ""<số tiền khách trả THỰC TẾ. Loại A: số cuối cùng sau chiết khấu ('Tổng thanh toán'/'Tiền thu'). Loại D: số trong 'THU X,XXX + SHIP' (chỉ lấy số trước '+SHIP'). Loại B/C: thường không có — để trống. ĐƠN VỊ NGHÌN ĐỒNG, chỉ ghi số. VD: 7280 nếu ghi '7,280'. KHÔNG THU SHIP / KO THU SHIP → tien_thu='0'>"",
  ""tien_ship"": ""<phí vận chuyển / tiền ship, ĐƠN VỊ NGHÌN ĐỒNG, chỉ ghi số. VD: 25. 'THU SHIP' hoặc 'THU X+SHIP' → có thu ship (tra bảng). Nếu không có ghi 0>"",
  ""ngay_lay"": ""<ngày lấy/giao hàng, format dd-MM-yyyy. Nhãn ship thường không có → để trống>"",
  ""invoice_type"": ""<loại đơn: 'COD' nếu có tiền thu bình thường; 'SHIP_ONLY_FREE' nếu thấy 'KHÔNG THU SHIP'/'KO THU SHIP'/'KHONG THU SHIP'/'KHÔNG THU'; 'SHIP_ONLY_PAID' nếu thấy 'THU SHIP' (không kèm 'KHÔNG'); mặc định 'COD'>""
}

QUY TẮC QUAN TRỌNG:
1. Địa chỉ GIAO HÀNG (của khách) — KHÔNG lấy địa chỉ cửa hàng (CN1/CN2/HOTLINE/chi nhánh).
   - Loại A: dòng 'Địa chỉ:' phía dưới tên khách (không phải dòng 'Địa chỉ:' đầu tiên của shop).
   - Loại B: địa chỉ trong sticker trắng (VD: '275/14B1 Đặng Nguyễn Cẩn, Phường Phú Lâm, TP. Hồ Chí Minh').
   - Loại C: địa chỉ trong sticker (VD: 'Tòa S501, vinhome grand park, Long Thạnh Mỹ, Thủ Đức').
   - Loại D: toàn bộ dòng địa chỉ dưới 'Khách hàng:' (VD: '401 Quang trung f10 gò vấp').
2. dia_chi = tất cả phần địa chỉ TRƯỚC phường/quận. Ví dụ:
   - '275/14B1 Đặng Nguyễn Cẩn, Phường Phú Lâm' → dia_chi='275/14B1 Đặng Nguyễn Cẩn', phuong='Phú Lâm', quan='6'
   - 'Tòa S501, vinhome grand park, Long Thạnh Mỹ, Thủ Đức' → dia_chi='Tòa S501, Vinhome Grand Park', phuong='Long Thạnh Mỹ', quan='Thủ Đức'
   - '401 Quang trung f10 gò vấp' → dia_chi='401 Quang Trung', phuong='Phường 10', quan='Gò Vấp'
   - '68 Nguyễn Trãi, P.3, Q.5' → dia_chi='68 Nguyễn Trãi', phuong='Phường 3', quan='5'
   - 'Block D2, Chung cư Sài Gòn Riverside, 4 Đào Trí, Q7' → dia_chi='Block D2, Chung cư Sài Gòn Riverside, 4 Đào Trí', quan='7'
   - 'Landmark 5 . Vinhome central park . F22 . Q.bthanh' → dia_chi='Landmark 5, Vinhome Central Park, F22', quan='Bình Thạnh'
3. QUẬN — nhận dạng mọi cách viết tắt thường gặp trên hóa đơn/nhãn viết tay/OCR:
   - Số: Q.1, Q1, q10, Q.10, Quận 5 → ghi: 1, 10, 5
   - Bình Thạnh: Q.bthanh, bthanh, BThạnh, Bình Thạnh → ghi: Bình Thạnh
   - Gò Vấp: Q.gvap, GVấp, Gò Vấp, gò vấp → ghi: Gò Vấp
   - Tân Phú: Q.tphu, TPhu, Tân Phú → ghi: Tân Phú
   - Tân Bình: Q.tbinh, TBình, Tân Bình → ghi: Tân Bình
   - Phú Nhuận: Q.pnhuan, PNhuận, Phú Nhuận → ghi: Phú Nhuận
   - Thủ Đức: Q.tduc, TĐức, Thủ Đức, thủ đức → ghi: Thủ Đức
   - Bình Tân: Q.btan, BTân, Bình Tân → ghi: Bình Tân
   - Quận 6: khu vực có Phường Phú Lâm → ghi: 6
4. TP.HCM năm 2025 đổi tên phường — nếu địa chỉ có tên phường/khu vực mới hãy suy ra đúng TÊN QUẬN:
   - An Hội Tây, Thông Tây Hội, An Nhơn → Gò Vấp
   - Phường 22, Phường 25, Phường 26, Phường 27, Phường 28 → Bình Thạnh
   - Long Thạnh Mỹ → Thủ Đức
   - Phú Lâm → Quận 6
5. Đọc loại đơn từ text trên nhãn:
   - 'KHÔNG THU SHIP' / 'KO THU SHIP' / 'KHÔNG THU' → invoice_type='SHIP_ONLY_FREE', tien_thu='0'
   - 'THU SHIP' (không kèm 'không/ko') → invoice_type='SHIP_ONLY_PAID', tien_thu='0'
   - 'THU X,XXX + SHIP' → invoice_type='COD', tien_thu = chỉ số X,XXX (trước '+SHIP'), ĐƠN VỊ NGHÌN
6. Với Loại B/C (nhãn ship): HOTLINE trên nhãn là SĐT của shop gửi hàng, KHÔNG phải SĐT khách. SĐT của khách là số kèm tên khách trong sticker địa chỉ.
7. Nếu không đọc được field nào thì để chuỗi rỗng """".
Chỉ trả về JSON, không có text khác.";

        public GeminiService(string apiKey)
        {
            _apiKey = apiKey;
        }

        public bool IsConfigured => !string.IsNullOrWhiteSpace(_apiKey);

        /// <summary>
        /// Gọi Gemini Vision để extract toàn bộ fields từ ảnh hóa đơn.
        /// Tự động thử lần lượt các model trong MODEL_FALLBACK_LIST cho đến khi thành công.
        /// Trả về (GeminiInvoiceResult, errorMessage) — errorMessage != "" khi tất cả model đều thất bại.
        /// </summary>
        public async Task<(GeminiInvoiceResult result, string error)> ParseInvoiceFromImageAsync(
            string imagePath
        )
        {
            if (!IsConfigured)
                return (null, "API key chưa cấu hình");
            if (!File.Exists(imagePath))
                return (null, $"File không tồn tại: {imagePath}");

            var lastError = "";
            foreach (var modelName in MODEL_FALLBACK_LIST)
            {
                try
                {
                    var googleAI = new GoogleAI(apiKey: _apiKey);
                    var model = googleAI.GenerativeModel(model: modelName);

                    var request = new GenerateContentRequest(PROMPT);
                    await request.AddMedia(imagePath);

                    var response = await model.GenerateContent(request);
                    var raw = response.Text?.Trim() ?? "";

                    System.Diagnostics.Debug.WriteLine(
                        $"[Gemini/{modelName}] Response for {Path.GetFileName(imagePath)}: {raw}"
                    );

                    if (string.IsNullOrWhiteSpace(raw))
                    {
                        lastError = $"[{modelName}] Response rỗng (quota hết hoặc safety filter)";
                        continue; // thử model tiếp theo
                    }

                    var parsed = ParseGeminiResponse(raw);
                    if (parsed == null)
                    {
                        lastError =
                            $"[{modelName}] Parse thất bại — raw: {raw.Substring(0, Math.Min(200, raw.Length))}";
                        continue; // thử model tiếp theo
                    }

                    System.Diagnostics.Debug.WriteLine($"[Gemini] OK với model: {modelName}");
                    return (parsed, "");
                }
                catch (Exception ex)
                {
                    string errDetail =
                        ex.InnerException != null
                            ? $"{ex.GetType().Name}: {ex.Message} | Inner: {ex.InnerException.Message}"
                            : $"{ex.GetType().Name}: {ex.Message}";

                    // Nếu là quota/rate limit/server overload → thử model tiếp theo
                    bool isRetryable =
                        errDetail.Contains("429")
                        || errDetail.Contains("TooManyRequests")
                        || errDetail.Contains("RESOURCE_EXHAUSTED")
                        || errDetail.Contains("quota")
                        || errDetail.Contains("503")
                        || errDetail.Contains("ServiceUnavailable")
                        || errDetail.Contains("UNAVAILABLE")
                        || errDetail.Contains("high demand")
                        || errDetail.Contains("try again");
                    lastError = $"[{modelName}] {errDetail}";
                    System.Diagnostics.Debug.WriteLine($"[Gemini/{modelName}] Error: {errDetail}");

                    if (isRetryable)
                        continue; // quota/overload → thử model kế tiếp
                    else
                        return (null, lastError); // lỗi khác (network, auth...) → báo ngay
                }
            }

            return (null, lastError);
        }

        /// <summary>
        /// Parse JSON response từ Gemini thành GeminiInvoiceResult.
        /// Dùng regex thay vì System.Text.Json để tránh dependency phức tạp.
        /// </summary>
        private static GeminiInvoiceResult ParseGeminiResponse(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return null;

            // Strip markdown code block nếu có: ```json ... ```
            json = Regex.Replace(json, @"^```[a-z]*\s*", "", RegexOptions.Multiline).Trim();
            json = Regex.Replace(json, @"```\s*$", "", RegexOptions.Multiline).Trim();

            string Extract(string field)
            {
                var m = Regex.Match(
                    json,
                    $@"""{field}""\s*:\s*""([^""]*?)""",
                    RegexOptions.IgnoreCase
                );
                return m.Success ? m.Groups[1].Value.Trim() : "";
            }

            var result = new GeminiInvoiceResult
            {
                TenShop = Extract("ten_shop"),
                TenKH = Extract("ten_kh"),
                Ma = Extract("ma"),
                DiaChi = Extract("dia_chi"),
                Phuong = NormalizeField(Extract("phuong")),
                Quan = NormalizeQuan(Extract("quan")),
                TienThu = NormalizeAmount(Extract("tien_thu")),
                TienShip = NormalizeAmount(Extract("tien_ship")),
                NgayLay = NormalizeDate(Extract("ngay_lay")),
                InvoiceType = NormalizeInvoiceType(Extract("invoice_type")),
            };

            // Nếu Gemini không trả về gì có ích thì trả null
            bool hasAny =
                !string.IsNullOrEmpty(result.TenKH)
                || !string.IsNullOrEmpty(result.Ma)
                || !string.IsNullOrEmpty(result.Quan)
                || !string.IsNullOrEmpty(result.DiaChi)
                || !string.IsNullOrEmpty(result.TienThu);
            return hasAny ? result : null;
        }

        // ─── Normalize helpers ────────────────────────────────────────────────

        /// <summary>
        /// Normalize quận về dạng AddressParser dùng:
        ///   "Quận 1" / "Q.1" / "1" → "1"
        ///   "Bình Thạnh" / "Q. Bình Thạnh" → "binh thanh"
        /// </summary>
        private static string NormalizeQuan(string quan)
        {
            if (string.IsNullOrWhiteSpace(quan))
                return "";
            quan = quan.Trim();
            // Bỏ prefix "Quận ", "Q.", "Q "
            quan = Regex
                .Replace(quan, @"^(?:qu[aâậ]n|q)\.?\s*", "", RegexOptions.IgnoreCase)
                .Trim();
            // Số thuần → giữ nguyên (VD: "1", "10", "12")
            if (Regex.IsMatch(quan, @"^\d{1,2}$"))
                return quan;
            // Tên quận → remove diacritics + lowercase
            return RemoveDiacritics(quan).ToLowerInvariant();
        }

        private static string NormalizeField(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return "";
            return RemoveDiacritics(s.Trim()).ToLowerInvariant();
        }

        /// <summary>
        /// Normalize số tiền từ Gemini về đơn vị nghìn đồng (khớp format app).
        /// Prompt đã yêu cầu Gemini trả đơn vị nghìn, nhưng guard thêm:
        ///   nếu Gemini vẫn trả số đầy đủ (≥ 10000) thì tự chia 1000.
        /// VD: "666" → "666" | "666000" → "666" | "0" / "" → ""
        /// </summary>
        private static string NormalizeAmount(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return "";
            // Bỏ ký tự không phải số
            var digits = Regex.Replace(s, @"[^\d]", "");
            if (digits == "" || digits == "0")
                return "";
            if (long.TryParse(digits, out long val))
            {
                // Nếu số ≥ 10000 → Gemini trả số đầy đủ → chia 1000
                if (val >= 10000)
                    return (val / 1000).ToString();
                return val.ToString();
            }
            return digits;
        }

        /// <summary>
        /// Normalize ngày về format dd/MM/yyyy.
        /// Gemini có thể trả về dd-MM-yyyy hoặc yyyy-MM-dd.
        /// </summary>
        private static string NormalizeDate(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return "";
            s = s.Trim();
            // Thay dấu - thành /
            s = s.Replace("-", "/");
            // yyyy/MM/dd → dd/MM/yyyy
            var ymd = Regex.Match(s, @"^(\d{4})/(\d{1,2})/(\d{1,2})$");
            if (ymd.Success)
                s = $"{ymd.Groups[3].Value}/{ymd.Groups[2].Value}/{ymd.Groups[1].Value}";
            return s;
        }

        /// <summary>
        /// Normalize invoice_type từ Gemini về một trong 3 giá trị chuẩn:
        ///   "SHIP_ONLY_FREE" | "SHIP_ONLY_PAID" | "COD"
        /// Trả về "" nếu Gemini không cung cấp / không nhận ra → caller giữ nguyên giá trị OCR.
        /// </summary>
        private static string NormalizeInvoiceType(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return "";
            var up = raw.Trim().ToUpperInvariant();
            if (up.Contains("FREE") || up.Contains("SHIP_ONLY_FREE"))
                return "SHIP_ONLY_FREE";
            if (up.Contains("PAID") || up.Contains("SHIP_ONLY_PAID"))
                return "SHIP_ONLY_PAID";
            if (up == "COD")
                return "COD";
            return "";
        }

        private static string RemoveDiacritics(string s)
        {
            if (string.IsNullOrEmpty(s))
                return s;
            var norm = s.Normalize(System.Text.NormalizationForm.FormD);
            var sb = new System.Text.StringBuilder();
            foreach (char c in norm)
                if (
                    System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c)
                    != System.Globalization.UnicodeCategory.NonSpacingMark
                )
                    sb.Append(c);
            return sb.ToString().Normalize(System.Text.NormalizationForm.FormC);
        }
    }
}

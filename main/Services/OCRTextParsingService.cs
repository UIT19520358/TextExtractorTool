using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace TextInputter.Services
{
    /// <summary>
    /// Service phân tích OCR text và extract các fields cần thiết.
    ///
    /// ⚠️ HARDCODED (cần discuss để cải thiện sau):
    ///   - Tất cả regex pattern đang viết cứng để khớp với hóa đơn của khách hiện tại.
    ///   - Các keyword: "HĐ", "Số HĐ", "shop", "cửa hàng", "mã", "địa chỉ", "tiền thu",
    ///     "tiền ship", "giảm", "chiết khấu" — phụ thuộc format hóa đơn.
    ///   - ExtractAllFields yêu cầu đúng 12 fields; nếu đổi template hóa đơn → cần update.
    /// </summary>
    public class OCRTextParsingService
    {
        // Gemini fallback — chỉ active khi GEMINI_API_KEY được điền trong AppConstants
        private readonly GeminiService _gemini = new GeminiService(AppConstants.GEMINI_API_KEY);

        // Path ảnh gốc hiện tại — được set bởi caller trước khi gọi ExtractAllFields
        // Dùng để Gemini đọc ảnh khi cần fallback
        public string CurrentImagePath { get; set; } = "";

        /// <summary>
        /// Extract tất cả 12 fields bắt buộc từ OCR text.
        /// Trả về danh sách các field bị thiếu (empty list = đủ hết).
        /// geminiLog (optional): nếu truyền vào, các dòng log Gemini sẽ được append vào list này
        /// thay vì tự ghi file — để caller gom cùng raw OCR + mapping vào 1 log thống nhất.
        /// </summary>
        public List<string> ExtractAllFields(
            string text,
            out Dictionary<string, string> fields,
            List<string> geminiLog = null
        )
        {
            fields = new Dictionary<string, string>();
            var missingFields = new List<string>();

            // 1. SHOP — lấy dòng bắt đầu "ĐOÀN" nhưng không phải dòng footer (đổi size, ngày kể...)
            // ⚠️ HARDCODED: tên shop cố định theo hóa đơn Đoàn Ngân Châu
            var shopLine = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None)
                .FirstOrDefault(l =>
                    (
                        l.Trim().StartsWith("ĐOÀN", StringComparison.OrdinalIgnoreCase)
                        || l.Trim().StartsWith("Đoàn", StringComparison.OrdinalIgnoreCase)
                    )
                    && !Regex.IsMatch(
                        l,
                        @"nhận đổi|sản phẩm|ngày kể|giặt ủi|khui hàng",
                        RegexOptions.IgnoreCase
                    )
                );
            if (shopLine == null)
            {
                var shopM = Regex.Match(
                    text,
                    @"(?:shop|cửa hàng|store)\s*[:=]?\s*([^\n\r,]+)",
                    RegexOptions.IgnoreCase
                );
                if (shopM.Success)
                {
                    var shopCandidate = shopM.Groups[1].Value.Trim();
                    // Loại bỏ nếu là dòng footer hoặc ghi chú thủ công
                    if (
                        !Regex.IsMatch(
                            shopCandidate,
                            AppConstants.SHOP_EXCLUSION_PATTERN,
                            RegexOptions.IgnoreCase
                        )
                    )
                        fields["SHOP"] = shopCandidate;
                    else
                        fields["SHOP"] = string.Empty;
                }
                else
                {
                    fields["SHOP"] = string.Empty;
                }
            }
            else
            {
                fields["SHOP"] = shopLine.Trim();
            }
            if (string.IsNullOrEmpty(fields["SHOP"]))
                missingFields.Add("SHOP");

            // 2. TÊN KH — "Khách hàng: ..." / "Khách hàng. ..." / "Khách hàng; ..." (OCR hay đọc ":" thành "."/";")
            var khMatch = Regex.Match(
                text,
                @"Kh[aá]ch\s*h[aà]ng\s*[:.=;]?\s*([^\n\r]+)",
                RegexOptions.IgnoreCase
            );
            var khRaw = khMatch.Success ? khMatch.Groups[1].Value.Trim() : string.Empty;
            // Strip các ký tự nhiễu đầu: ". ", "; ", "VIP-", "VIP " ...
            khRaw = Regex.Replace(khRaw, @"^[\s;.,\-–—:]+", "").Trim();
            // Nếu có "VIP" ở đầu (như "VIP- TÊN KH") thì bỏ
            khRaw = Regex.Replace(khRaw, @"^VIP[\s\-–—]*", "", RegexOptions.IgnoreCase).Trim();
            // Loại C — nhãn viết tay: không có "Khách hàng:" nhưng có SĐT kèm tên khách
            // VD: "0906590639 THƯƠNG 1994" → lấy phần tên sau SĐT
            if (string.IsNullOrEmpty(khRaw))
            {
                var soPhoiTenMatch = Regex.Match(
                    text,
                    @"(?:^|\n)\s*(0\d{9})\s+([A-ZĐÀÁẢÃẠĂẮẶẰẲẴÂẤẦẨẪẬÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴ][^\n\r]{1,40})",
                    RegexOptions.IgnoreCase
                );
                if (soPhoiTenMatch.Success)
                    khRaw = soPhoiTenMatch.Groups[2].Value.Trim();
            }
            fields["TÊN KH"] = khRaw;
            if (string.IsNullOrEmpty(fields["TÊN KH"]))
                missingFields.Add("TÊN KH");

            // 3. MÃ (Số HĐ) — OCR hay đọc méo "So HD:" thành nhiều dạng khác nhau
            // ⚠️ HARDCODED: format "HD\d+" theo template hóa đơn Đoàn Ngân Châu
            var maField = Regex.Match(
                text,
                @"S[oố]\s*H[ĐD\d]\s*[:=]?\s*(HD\d+)",
                RegexOptions.IgnoreCase
            );
            if (!maField.Success)
                maField = Regex.Match(
                    text,
                    @"(?:Số\s*HĐ|HĐ\s*số|Invoice\s*No)\s*[:=]?\s*([^\n\r\s,]+)",
                    RegexOptions.IgnoreCase
                );
            if (!maField.Success)
                // Fallback: tìm "HD" theo sau ngay bởi số — xuất hiện bất kỳ đâu trong dòng
                maField = Regex.Match(text, @"\b(HD\d{4,})\b", RegexOptions.IgnoreCase);
            fields["MÃ"] = maField.Success
                ? maField.Groups[1].Value.Trim().ToUpper()
                : string.Empty;
            if (string.IsNullOrEmpty(fields["MÃ"]))
                missingFields.Add("MÃ");

            // 4–7. Address parts — parse địa chỉ thành ĐỊA CHỈ (pre-quận) + QUẬN
            string addressLine = ExtractAddressLine(text);
            // Parse QUẬN từ raw addressLine (còn đủ quận) TRƯỚC khi strip
            var parsed = AddressParser.Parse(addressLine);
            fields["QUẬN"] = parsed.Quan;

            // Fallback QUẬN: nếu AddressParser không ra → scan toàn bộ raw OCR text tìm "Quận X"
            // Xử lý trường hợp OCR wrap dòng giữa tên quận (VD: "Quận B\nh Thạnh" → "Bình Thạnh")
            if (string.IsNullOrEmpty(fields["QUẬN"]))
                fields["QUẬN"] = ExtractDistrictFromRawText(text);

            // ĐỊA CHỈ = strip quận (và phường) ra khỏi addressLine — quận đã có cột riêng
            fields["ĐỊA CHỈ"] = StripDistrictAndWard(addressLine);

            // 8. TIỀN THU — Ưu tiên lấy số tiền CUỐI CÙNG trong text (sau chiết khấu)
            // Lý do: hóa đơn có chiết khấu → "Tổng tiền hàng" ≠ "Tổng thanh toán",
            //        số cuối cùng in trên hóa đơn là số khách thực trả.
            // Fallback theo thứ tự:
            //   0. Loại D — nhãn "THU X,XXX + SHIP": chỉ lấy số trước "+SHIP"
            //   1. Số tiền cuối cùng trong raw text (>= 1k, định dạng X,XXX hoặc XXX,XXX,XXX)
            //   2. Keyword "tổng thanh toán" / "tiền thu" / "total" (nếu không tìm được số cuối)
            //   3. Keyword "t.tiên" / "T.Tiền"
            {
                // Bước 0: Loại D — "THU 7,280 + SHIP" / "THU 7280+SHIP" / "THU: 7,280 +SHIP"
                // Chỉ lấy số trước "+SHIP"; phần "+SHIP" nghĩa là có thu ship → giữ COD logic
                // ⚠️ Nhãn viết tay luôn ghi số ở đơn vị NGHÌN ĐỒNG (VD: "7.280" = 7280k)
                //    → CHỈ strip dấu ngăn cách, KHÔNG chia 1000 như NormalizeToThousands
                var thuPlusShip = Regex.Match(
                    text,
                    @"\bTHU\s*[:\s]*(\d{1,3}(?:[.,]\d{3})+|\d{3,})\s*\+\s*SHIP",
                    RegexOptions.IgnoreCase
                );
                if (thuPlusShip.Success)
                {
                    // Strip dấu phẩy/chấm ngăn cách nghìn, giữ nguyên giá trị (đã là nghìn đồng)
                    var raw0 = thuPlusShip.Groups[1].Value.TrimEnd('.', ',');
                    var digits0 = raw0.Replace(",", "").Replace(".", "");
                    fields["TIỀN THU"] = long.TryParse(digits0, out long v0)
                        ? v0.ToString()
                        : digits0;
                    // "THU X + SHIP" = COD (thu tiền hàng VÀ thu ship) → giữ COD, không đổi type
                }

                // Bước 1: lấy số cuối cùng trong text (>= 3 chữ số, có dấu phẩy ngăn cách)
                // Pattern: số có định dạng tiền VD: "666,000" / "1,200,000" / "666.000"
                if (string.IsNullOrEmpty(fields.GetValueOrDefault("TIỀN THU", "")))
                {
                    var allMoneyMatches = Regex.Matches(text, @"\b(\d{1,3}(?:[.,]\d{3})+)\b");
                    if (allMoneyMatches.Count > 0)
                    {
                        var lastRaw = allMoneyMatches[allMoneyMatches.Count - 1].Groups[1].Value;
                        fields["TIỀN THU"] = NormalizeToThousands(lastRaw.TrimEnd('.', ','));
                    }
                }

                // Bước 2 (fallback nếu không tìm được số cuối): keyword rõ ràng
                if (string.IsNullOrEmpty(fields.GetValueOrDefault("TIỀN THU", "")))
                    fields["TIỀN THU"] = ExtractAmountLine(
                        text,
                        @"tiền thu|thu tiền|tổng thanh toán|tong thanh toan|thanh toán|thanh toan|total"
                    );

                // Bước 3 (fallback): "t.tiên" / "T.Tiền"
                if (string.IsNullOrEmpty(fields.GetValueOrDefault("TIỀN THU", "")))
                    fields["TIỀN THU"] = ExtractAmountLine(text, @"t\.ti[eêề]n|T\.Ti[eêề]n");
            }

            // 9. TIỀN SHIP — optional: nếu không có trong ảnh thì để trống,
            // ProcessImages() sẽ tự tra bảng phí ship theo quận để điền vào.
            fields["TIỀN SHIP"] = ExtractAmountLine(text, "tiền ship|ship|vận chuyển");

            // Sanity check TIỀN SHIP: nếu bằng hoặc lớn hơn TIỀN THU → chắc chắn lấy nhầm số
            // (VD: OCR có chữ "Ship" ở footer, lấy nhầm số 820,000 của tiền hàng)
            // → reset về trống, ProcessImages() sẽ tra bảng phí ship theo quận
            if (
                !string.IsNullOrEmpty(fields["TIỀN SHIP"])
                && !string.IsNullOrEmpty(fields["TIỀN THU"])
                && long.TryParse(fields["TIỀN SHIP"], out long ship)
                && long.TryParse(fields["TIỀN THU"], out long thu)
                && ship >= thu
            )
                fields["TIỀN SHIP"] = "";

            // 9b. INVOICE TYPE — detect loại đơn từ keyword trong text:
            //   "không thu ship" / "ko thu ship" / "khong thu" → SHIP_ONLY_FREE  (thu=0, hàng=-ship)
            //   "thu ship" (không kèm "không") → SHIP_ONLY_PAID  (thu=0, hàng=+ship)
            //   còn lại (có TIỀN THU > 0) → COD (thu=x, hàng=x+ship)  ← format cũ
            bool hasKhongThuShip = Regex.IsMatch(
                text,
                @"không\s+thu\s+ship|ko\s+thu\s+ship|khong\s+thu\s+ship|không\s+thu\b|ko\s+thu\b",
                RegexOptions.IgnoreCase
            );
            bool hasThuShip =
                !hasKhongThuShip && Regex.IsMatch(text, @"\bthu\s+ship\b", RegexOptions.IgnoreCase);

            if (hasKhongThuShip)
            {
                fields["INVOICE_TYPE"] = "SHIP_ONLY_FREE"; // hàng = -ship
                fields["TIỀN THU"] = "0"; // không thu tiền khách
            }
            else if (hasThuShip)
            {
                fields["INVOICE_TYPE"] = "SHIP_ONLY_PAID"; // hàng = +ship
                fields["TIỀN THU"] = "0";
            }
            else
            {
                fields["INVOICE_TYPE"] = "COD"; // format cũ: hàng = thu + ship
            }

            // 10. NGÀY LẤY
            fields["NGÀY LẤY"] = ExtractDate(text);

            // ── GEMINI FALLBACK ──────────────────────────────────────────────────────
            // Sau khi OCR parsing xong toàn bộ → kiểm tra field nào còn trống.
            // Nếu bất kỳ field quan trọng nào thiếu → trigger Gemini đọc ảnh gốc.
            // Chỉ log FAILED nếu sau Gemini vẫn còn trống.
            bool needGemini =
                (
                    string.IsNullOrEmpty(fields["SHOP"])
                    || string.IsNullOrEmpty(fields["QUẬN"])
                    || string.IsNullOrEmpty(fields["ĐỊA CHỈ"])
                    || string.IsNullOrEmpty(fields["TIỀN THU"])
                    || string.IsNullOrEmpty(fields["TÊN KH"])
                    || string.IsNullOrEmpty(fields["MÃ"])
                    || string.IsNullOrEmpty(fields["NGÀY LẤY"])
                )
                && _gemini.IsConfigured
                && !string.IsNullOrEmpty(CurrentImagePath);
            if (needGemini)
            {
                string imgName = System.IO.Path.GetFileName(CurrentImagePath);
                // Log các field còn thiếu để dễ debug
                var missingBefore = new List<string>();
                if (string.IsNullOrEmpty(fields["SHOP"]))
                    missingBefore.Add("SHOP");
                if (string.IsNullOrEmpty(fields["QUẬN"]))
                    missingBefore.Add("QUẬN");
                if (string.IsNullOrEmpty(fields["ĐỊA CHỈ"]))
                    missingBefore.Add("ĐỊA CHỈ");
                if (string.IsNullOrEmpty(fields["TIỀN THU"]))
                    missingBefore.Add("TIỀN THU");
                if (string.IsNullOrEmpty(fields["TÊN KH"]))
                    missingBefore.Add("TÊN KH");
                if (string.IsNullOrEmpty(fields["MÃ"]))
                    missingBefore.Add("MÃ");
                if (string.IsNullOrEmpty(fields["NGÀY LẤY"]))
                    missingBefore.Add("NGÀY LẤY");
                AddGeminiLog(
                    geminiLog,
                    $"TRIGGERED for: {imgName} | THIẾU: {string.Join(", ", missingBefore)}"
                );
                System.Diagnostics.Debug.WriteLine(
                    $"[Gemini] Fallback triggered for: {imgName} | missing: {string.Join(", ", missingBefore)}"
                );

                var (g, geminiError) = _gemini
                    .ParseInvoiceFromImageAsync(CurrentImagePath)
                    .GetAwaiter()
                    .GetResult();
                if (g != null)
                {
                    // SHOP — Gemini đọc được tên shop từ header hóa đơn
                    if (string.IsNullOrEmpty(fields["SHOP"]) && !string.IsNullOrEmpty(g.TenShop))
                        fields["SHOP"] = g.TenShop;
                    // Địa chỉ
                    if (string.IsNullOrEmpty(fields["QUẬN"]) && !string.IsNullOrEmpty(g.Quan))
                    {
                        // Phòng Gemini trả tên phường thay vì tên quận → tra WARD_TO_DISTRICT_MAP
                        string quanValue = g.Quan;
                        var quanNorm = System
                            .Text.RegularExpressions.Regex.Replace(
                                g.Quan.Normalize(System.Text.NormalizationForm.FormD),
                                @"[^a-z0-9 ]",
                                "",
                                System.Text.RegularExpressions.RegexOptions.IgnoreCase
                            )
                            .ToLowerInvariant()
                            .Trim();
                        quanNorm = System.Text.RegularExpressions.Regex.Replace(
                            quanNorm,
                            @"\s+",
                            " "
                        );
                        if (
                            AppConstants.WARD_TO_DISTRICT_MAP.TryGetValue(
                                quanNorm,
                                out var mappedQuan
                            )
                        )
                            quanValue = mappedQuan;
                        fields["QUẬN"] = quanValue;
                    }
                    // SỐ NHÀ: override nếu đang là raw fallback (Gemini tách chính xác hơn)
                    if (
                        (string.IsNullOrEmpty(fields["ĐỊA CHỈ"])) && !string.IsNullOrEmpty(g.DiaChi)
                    )
                        fields["ĐỊA CHỈ"] = g.DiaChi;
                    if (string.IsNullOrEmpty(fields["TÊN KH"]) && !string.IsNullOrEmpty(g.TenKH))
                        fields["TÊN KH"] = g.TenKH;
                    if (string.IsNullOrEmpty(fields["MÃ"]) && !string.IsNullOrEmpty(g.Ma))
                        fields["MÃ"] = g.Ma;
                    if (
                        string.IsNullOrEmpty(fields["TIỀN THU"]) && !string.IsNullOrEmpty(g.TienThu)
                    )
                        fields["TIỀN THU"] = g.TienThu;
                    if (
                        string.IsNullOrEmpty(fields["TIỀN SHIP"])
                        && !string.IsNullOrEmpty(g.TienShip)
                    )
                        fields["TIỀN SHIP"] = g.TienShip;
                    if (
                        string.IsNullOrEmpty(fields["NGÀY LẤY"]) && !string.IsNullOrEmpty(g.NgayLay)
                    )
                        fields["NGÀY LẤY"] = g.NgayLay;

                    // INVOICE_TYPE — Gemini detect loại đơn từ nhãn (KHÔNG THU SHIP / THU SHIP / THU X+SHIP)
                    // Chỉ override nếu OCR rule-based chưa detect được (vẫn là "COD" mặc định)
                    // hoặc nếu Gemini trả về type khác hẳn (ưu tiên Gemini cho các nhãn mới)
                    if (!string.IsNullOrEmpty(g.InvoiceType))
                    {
                        string currentType = fields.GetValueOrDefault("INVOICE_TYPE", "COD");
                        // Nếu OCR vẫn là "COD" (chưa detect đặc biệt) → tin Gemini
                        // Nếu Gemini và OCR đồng ý hoặc Gemini cụ thể hơn → dùng Gemini
                        if (currentType == "COD" || g.InvoiceType != "COD")
                        {
                            fields["INVOICE_TYPE"] = g.InvoiceType;
                            // Sync TIỀN THU theo loại đơn mà Gemini phát hiện
                            if (
                                g.InvoiceType == "SHIP_ONLY_FREE"
                                || g.InvoiceType == "SHIP_ONLY_PAID"
                            )
                                fields["TIỀN THU"] = "0";
                            AddGeminiLog(
                                geminiLog,
                                $"INVOICE_TYPE override → {g.InvoiceType} (OCR was: {currentType})"
                            );
                        }
                    }

                    // PHƯỜNG — Gemini có thể tách được phường khi OCR không làm được
                    if (
                        string.IsNullOrEmpty(fields.GetValueOrDefault("PHƯỜNG", ""))
                        && !string.IsNullOrEmpty(g.Phuong)
                    )
                        fields["PHƯỜNG"] = g.Phuong;

                    string resultLine =
                        $"OK | SHOP={g.TenShop} | QUẬN={g.Quan} | TÊN KH={g.TenKH} | MÃ={g.Ma}"
                        + $" | THU={g.TienThu} | SHIP={g.TienShip} | NGÀY={g.NgayLay} | TYPE={g.InvoiceType}";
                    AddGeminiLog(geminiLog, resultLine);
                    System.Diagnostics.Debug.WriteLine($"[Gemini] {resultLine}");
                }
                else
                {
                    string failMsg = string.IsNullOrEmpty(geminiError)
                        ? "FAILED — response null/empty"
                        : $"FAILED — {geminiError}";
                    AddGeminiLog(geminiLog, failMsg);
                    System.Diagnostics.Debug.WriteLine($"[Gemini] {failMsg}");
                }
            }

            // ── MISSING FIELDS (sau cả OCR + Gemini) ────────────────────────────────
            // SHOP final fallback: nếu vẫn rỗng sau tất cả parsing + Gemini → dùng tên shop mặc định
            if (string.IsNullOrWhiteSpace(fields.GetValueOrDefault("SHOP", "")))
            {
                fields["SHOP"] = AppConstants.SHOP_DEFAULT;
                missingFields.Remove("SHOP"); // không còn missing nữa
            }

            // NGÀY LẤY final fallback: nếu vẫn rỗng → dùng ngày hôm nay
            if (string.IsNullOrEmpty(fields["NGÀY LẤY"]))
            {
                fields["NGÀY LẤY"] = DateTime.Today.ToString(AppConstants.DATE_FORMAT_EXCEL);
                missingFields.Remove("NGÀY LẤY"); // không còn missing nữa
            }

            // Chỉ log FAILED cho field nào vẫn còn trống sau khi Gemini đã cố fill
            if (string.IsNullOrEmpty(fields["ĐỊA CHỈ"]))
                missingFields.Add("ĐỊA CHỈ");
            if (string.IsNullOrEmpty(fields["QUẬN"]))
                missingFields.Add("QUẬN");
            // Đơn ship-only (SHIP_ONLY_FREE / SHIP_ONLY_PAID): TIỀN THU = 0 là hợp lệ, không báo thiếu
            string invoiceType = fields.GetValueOrDefault("INVOICE_TYPE", "COD");
            bool isShipOnly = invoiceType == "SHIP_ONLY_FREE" || invoiceType == "SHIP_ONLY_PAID";
            if (!isShipOnly && string.IsNullOrEmpty(fields["TIỀN THU"]))
                missingFields.Add("TIỀN THU");
            // NGÀY LẤY luôn có giá trị (fallback về today ở trên) → không cần check thiếu

            // NGƯỜI ĐI / NGƯỜI LẤY: được nhập thủ công từ UI, không extract ở đây
            fields["NGƯỜI ĐI"] = "";
            fields["NGƯỜI LẤY"] = "";

            return missingFields;
        }

        // ─── Private helpers ───────────────────────────────────────────────────

        /// <summary>
        /// Fallback: scan toàn bộ raw OCR text để tìm tên quận dạng "Quận X" / "Q. X".
        /// Ghép text đa dòng trước khi match — xử lý OCR wrap giữa tên quận
        /// (VD: "Quận B\nh Thạnh -" → nhận ra "Bình Thạnh").
        /// Trả về giá trị đã normalize (khớp DistrictDict của AddressParser).
        /// </summary>
        private string ExtractDistrictFromRawText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            // Thử 2 cách ghép dòng:
            //   joinChar="" → xóa \n không có space → bắt wrap giữa từ: "Quận B\nh Thạnh" → "Quận Bh Thạnh"
            //   joinChar=" " → thay bằng space → bắt wrap đúng ranh giới từ bình thường
            foreach (var joinChar in new[] { "", " " })
            {
                var flat = Regex.Replace(text, @"[\r\n]+", joinChar);

                // Match "Quận <tên>" hoặc "Q. <tên>" — tên có thể là số hoặc chữ (≤4 từ)
                // Dừng lại trước dấu "-", "–", "," hoặc hết chuỗi
                var m = Regex.Match(
                    flat,
                    @"(?:Qu[aâậ]n|Q\.)\s*([A-ZĐÀÁẢÃẠĂẮẶẰẲẴÂẤẦẨẪẬÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴ\d][^\-–,\n]{0,40}?)(?:\s*[-–,]|\s{2,}|$)",
                    RegexOptions.IgnoreCase
                );
                if (!m.Success)
                    continue;

                var raw = m.Groups[1].Value.Trim().TrimEnd('-', '–', ' ');
                if (string.IsNullOrEmpty(raw))
                    continue;

                // Bỏ prefix "Quận " / "Q." nếu còn sót
                raw = Regex
                    .Replace(raw, @"^(?:Qu[aâậ]n|Q)\.?\s*", "", RegexOptions.IgnoreCase)
                    .Trim();

                // Nếu là số thuần → trả thẳng
                if (Regex.IsMatch(raw, @"^\d{1,2}$"))
                    return raw;

                // Normalize về không dấu + lowercase để AddressParser nhận ra
                var norm = RemoveDiacritics(raw).ToLowerInvariant().Trim();
                // Bỏ các từ trailing không phải tên quận (OCR noise)
                norm = Regex.Replace(norm, @"\s+[-–]\s*$", "").Trim();

                // Tra DistrictDict thông qua AddressParser (reuse logic normalize sẵn)
                // Dùng Parse() với chuỗi "q.<raw>" để tận dụng toàn bộ lookup
                var probe = AddressParser.Parse("q. " + raw);
                if (!string.IsNullOrEmpty(probe.Quan))
                    return probe.Quan;

                // Nếu probe không ra nhưng norm hợp lệ → trả norm (vẫn tốt hơn để trống)
                if (!string.IsNullOrEmpty(norm))
                    return norm;
            }
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

        private string ExtractAddressLine(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";
            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            // Hóa đơn thật có 2 dòng "Địa Chi/Chỉ:":
            //   Dòng 1 — địa chỉ CỬA HÀNG (CN1-..., CN2-...)  ← bỏ qua
            //   Dòng 2 — địa chỉ KHÁCH HÀNG                   ← lấy cái này
            // → Lấy dòng CUỐI CÙNG có "địa chỉ" / "address", không phải dòng đầu tiên.
            string found = "";
            int foundLineIdx = -1;
            for (int idx = 0; idx < lines.Length; idx++)
            {
                var line = lines[idx];
                if (
                    line.IndexOf("địa chỉ", StringComparison.OrdinalIgnoreCase) >= 0
                    || line.IndexOf("địa chi", StringComparison.OrdinalIgnoreCase) >= 0
                    || line.IndexOf("dia chi", StringComparison.OrdinalIgnoreCase) >= 0
                    || line.IndexOf("address", StringComparison.OrdinalIgnoreCase) >= 0
                    // Loại B — nhãn có "Địa chỉ mới:" (địa chỉ giao mới/tạm thời)
                    || line.IndexOf("địa chỉ mới", StringComparison.OrdinalIgnoreCase) >= 0
                    || line.IndexOf("dia chi moi", StringComparison.OrdinalIgnoreCase) >= 0
                )
                {
                    int colon = line.IndexOf(':');
                    string candidate = colon >= 0 ? line.Substring(colon + 1).Trim() : line.Trim();

                    // Bỏ qua nếu là địa chỉ shop (chứa "CN1", "CN2", "HOTLINE", số điện thoại)
                    if (
                        Regex.IsMatch(
                            candidate,
                            @"CN\d|HOTLINE|CHUYÊN SỈ|\b09\d{8}\b|\b03\d{8}\b",
                            RegexOptions.IgnoreCase
                        )
                    )
                        continue;

                    // Bỏ qua nếu candidate quá ngắn hoặc chỉ là "chi" (OCR đọc nhầm label)
                    if (candidate.Length < 4)
                        continue;

                    found = candidate;
                    foundLineIdx = idx;
                    // Không break — tiếp tục để lấy dòng cuối cùng hợp lệ
                }
            }

            // Nếu dòng "Địa chỉ:" lấy được chỉ toàn ký tự OCR sai (không có chữ số và không có từ địa chỉ),
            // thử ghép với dòng kế tiếp.
            if (foundLineIdx >= 0 && !string.IsNullOrEmpty(found))
            {
                bool hasDigit = Regex.IsMatch(found, @"\d");
                bool hasVietWord = Regex.IsMatch(
                    found,
                    @"\b(đường|phường|quận|hẻm|ngõ|nguyễn|trần|lê|phú|bình|tân|thành|hồ|hưng|minh|cộng|hòa)\b",
                    RegexOptions.IgnoreCase
                );
                if (!hasDigit && !hasVietWord && foundLineIdx + 1 < lines.Length)
                {
                    string nextLine = lines[foundLineIdx + 1].Trim();
                    if (
                        nextLine.Length >= 5
                        && Regex.IsMatch(
                            nextLine,
                            @"\d|nguyễn|trần|lê|phường|quận|đường|hẻm|ngõ",
                            RegexOptions.IgnoreCase
                        )
                    )
                    {
                        found = nextLine;
                    }
                }

                // Ghép dòng kế tiếp nếu địa chỉ bị OCR wrap dở (dòng sau bắt đầu bằng phần cuối bị cắt)
                // VD: "Địa chỉ: cổng số 2 ... chung cư khang gia, đư"  ← bị cắt ở "đư"
                //     dòng tiếp:  "ờng 45, an hội tây, hcm"             ← phần còn lại
                // Dấu hiệu: found kết thúc bằng chữ thường lửng (không phải dấu phẩy/gạch/số)
                // và dòng tiếp bắt đầu bằng chữ thường (continuation)
                var wrapContinuation =
                    @"^[a-záàảãạăắặằẳẵâấầẩẫậéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵđ]";
                bool foundEndsMidWord2 = Regex.IsMatch(
                    found,
                    @"[a-zA-ZáàảãạăắặằẳẵâấầẩẫậéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵđĐ]$"
                );
                if (foundEndsMidWord2 && foundLineIdx + 1 < lines.Length)
                {
                    string nextLine2 = lines[foundLineIdx + 1].Trim();
                    bool nextStartsContinuation =
                        nextLine2.Length > 0 && Regex.IsMatch(nextLine2, wrapContinuation);
                    if (nextStartsContinuation)
                        found = found + nextLine2;
                }
                // Trường hợp OCR in phần tiếp theo lên dòng TRƯỚC dòng "Địa chỉ:" (OCR layout ngược)
                // VD: dòng(foundLineIdx-1) = "ờng 45, an hội tây, hcm"   ← continuation
                //     dòng(foundLineIdx)   = "Địa chỉ: cổng số 2 ... đư"  ← bị cắt ở "đư"
                // Dấu hiệu: found kết thúc mid-word VÀ dòng TRƯỚC bắt đầu bằng chữ thường
                if (foundEndsMidWord2 && foundLineIdx - 1 >= 0)
                {
                    string prevLine = lines[foundLineIdx - 1].Trim();
                    bool prevStartsContinuation =
                        prevLine.Length > 0 && Regex.IsMatch(prevLine, wrapContinuation);
                    if (prevStartsContinuation)
                        found = found + prevLine;
                }
            }

            // Strip trailing garbage: dấu "-", "ạ", "…", khoảng trắng thừa
            found = Regex.Replace(found, @"[\s\-–—ạ\.…]+$", "").Trim();

            // Strip "TP HCM", "TP. HCM", "Hồ Chí Minh", "Hồ Chí Minh" khỏi cuối địa chỉ
            found = Regex
                .Replace(
                    found,
                    @",?\s*(?:TP\.?\s*H[CG]M|Hồ\s*Chí\s*Minh|HCM|TP\.?\s*HCM)\s*[ạa]?$",
                    "",
                    RegexOptions.IgnoreCase
                )
                .Trim();

            // Strip "Phường <tên>" / "P. <tên>" ở cuối (phường tên — không số, không cần map)
            // VD: "11 In Dung Vương Phường An Đông" → "11 In Dung Vương"
            found = Regex
                .Replace(
                    found,
                    @",?\s*Ph[uướừửữ][oôờ]ng\s+(?:[A-ZĐÀÁẢÃẠĂẮẶẰẲẴÂẤẦẨẪẬ][^\n,]*|An Đông|Tân Sơn Nhì|[^\d,]+)\s*$",
                    "",
                    RegexOptions.IgnoreCase
                )
                .Trim();

            // NOTE: quận KHÔNG strip ở đây — ExtractAddressLine trả về raw (kể cả quận)
            // Quận được parse riêng bởi AddressParser.Parse(), sau đó StripDistrictAndWard()
            // sẽ strip quận ra khỏi chuỗi để gán vào fields["ĐỊA CHỈ"].

            // Strip prefix "Đc:", "Dc:", "DC:" đầu địa chỉ (VD: "Dc: Số 1 Đinh Lễ...")
            found = Regex.Replace(found, @"^[Đđ][Cc]\s*:?\s*", "").Trim();

            // Chuẩn hóa "pXqY" / "p.X.qY" thành "pX, qY" để AddressParser split đúng
            // VD: "p10q10" → "p10, q10" | "p7.q5" → "p7, q5"
            found = Regex.Replace(
                found,
                @"\b(p\.?\d{1,2})\s*\.?\s*(q\.?\d{1,2})\b",
                "$1, $2",
                RegexOptions.IgnoreCase
            );

            // Strip dấu ":" sau số nhà (VD: "2181: Ng vẫn cư" → "2181 Ng vẫn cư")
            found = Regex.Replace(found, @"^(\d+)\s*:\s*", "$1 ").Trim();

            // Strip rác sau dấu ". " khi theo sau là chữ hoa (VD: ". Tiệm nail thuỳ thỏ")
            // NGOẠI LỆ 1: "Đ." / "đ." là viết tắt hợp lệ của "Đường" — KHÔNG strip
            // NGOẠI LỆ 2: phần sau dấu "." chứa keyword địa chỉ (Q., quận, phường, đường, hẻm, số)
            //   VD: "Landmark 5 . Vinhome central park. F22 . Q.bthanh" → KHÔNG strip (có Q.)
            //   VD: "363 Đ. Hùng Vương . Tiệm nail" → strip ". Tiệm nail"
            found = Regex
                .Replace(
                    found,
                    @"(?<![Đđ])\s*\.\s+(?![Qq]u?[aâậ]?\.|[Qq]u[aâậ]n|[Pp]h[uướ]|[Đđ][uưứ][oờô]ng|h[eẻ]m|ng[oõ]|\d)[A-ZĐÀÁẢÃẠĂẮẶẰẲẴÂẤẦẨẪẬ][^\n]*$",
                    ""
                )
                .Trim();

            // Strip trailing business name sau " - " hoặc " – " (VD: "363 Đ. Hùng Vương - Khải Nam Transpost")
            // Chỉ strip nếu phần sau " - " không phải keyword địa chỉ (đường/phường/quận/hẻm)
            found = Regex
                .Replace(
                    found,
                    @"\s+[-–—]\s+(?!đường|phường|quận|hẻm|ngõ|p\d|q\d)[^\d,]+$",
                    "",
                    RegexOptions.IgnoreCase
                )
                .Trim();

            // Strip trailing garbage lần 2: dấu "-", "–", khoảng trắng thừa còn sót
            found = Regex.Replace(found, @"[\s\-–—]+$", "").Trim();

            return found;
        }

        /// <summary>
        /// Strip quận và phường khỏi địa chỉ để lưu vào cột ĐỊA CHỈ (quận đã có cột riêng).
        /// Gọi SAU khi đã parse Quan từ raw addressLine bằng AddressParser.Parse().
        /// </summary>
        private static string StripDistrictAndWard(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
                return address;
            var s = address;

            // Strip phường dài: "Phường An Đông", "Phường 14"
            s = Regex
                .Replace(
                    s,
                    @",?\s*Ph[uướừửữ][oôờ]ng\s+(?:[A-ZĐÀÁẢÃẠĂẮẶẰẲẴÂẤẦẨẪẬ][^\n,]*|An Đông|Tân Sơn Nhì|[^\d,]+|\d{1,2})\s*$",
                    "",
                    RegexOptions.IgnoreCase
                )
                .Trim();

            // Strip phường viết tắt cuối: "p1", "p.1", "p 1", "p.14"
            s = Regex
                .Replace(s, @",?\s*\bp\.?\s*\d{1,2}\b\s*$", "", RegexOptions.IgnoreCase)
                .Trim();

            // Strip "Quận X" / "Q.X" bất kỳ nơi nào trong chuỗi (kể cả giữa — VD: "(22 Quận Bình Thạnh (...")
            // Bao gồm: số quận và tên quận chữ
            var districtNames =
                @"(?:Bình\s*Thạnh|Gò\s*Vấp|Thủ\s*Đức|Tân\s*Phú|Tân\s*Bình|Bình\s*Tân|Phú\s*Nhuận|Nhà\s*Bè|Hóc\s*Môn|Bình\s*Chánh|Cần\s*Giờ|Củ\s*Chi|\d{1,2})";
            // 1) Dạng có prefix "Quận"/"Q.": strip từ "Quận" đến hết tên (kể cả mọi vị trí)
            s = Regex
                .Replace(
                    s,
                    @"[,\s\(]*(?:Qu[aâậ]n|Q)\.?\s*" + districtNames + @"\s*",
                    " ",
                    RegexOptions.IgnoreCase
                )
                .Trim();

            // 2) Tên quận chữ bare ở cuối (không có prefix "Quận") — chỉ strip khi ở cuối $
            //    để tránh false positive giữa chuỗi (VD: "Bình Thạnh" là tên tòa nhà)
            s = Regex
                .Replace(
                    s,
                    @"[,\s]*(?:Bình\s*Thạnh|Gò\s*Vấp|Thủ\s*Đức|Tân\s*Phú|Tân\s*Bình|Bình\s*Tân|Phú\s*Nhuận|Nhà\s*Bè|Hóc\s*Môn|Bình\s*Chánh|Cần\s*Giờ|Củ\s*Chi)\s*$",
                    "",
                    RegexOptions.IgnoreCase
                )
                .Trim();

            // Strip trailing comma/dash/space/open-paren thừa
            s = Regex.Replace(s, @"[\s,\-–—\(]+$", "").Trim();
            return s;
        }

        private string ExtractAmountLine(string text, string keywords)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";
            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            for (int i = 0; i < lines.Length; i++)
            {
                if (!Regex.IsMatch(lines[i], keywords, RegexOptions.IgnoreCase))
                    continue;

                // Số cùng dòng với keyword
                // [.,] cho phép cả dấu phẩy lẫn dấu chấm làm ký tự ngăn cách nghìn
                var m = Regex.Match(lines[i], @"[\d][,.\d]*\d[,.]?");
                if (m.Success)
                    return NormalizeToThousands(m.Value.TrimEnd('.', ','));
                // Số ở dòng tiếp theo (VD: "Tổng thanh toán:\n592,000")
                if (i + 1 < lines.Length)
                {
                    var next = Regex.Match(lines[i + 1].Trim(), @"^[\d][,.\d]*\d[,.]?$");
                    if (next.Success)
                        return NormalizeToThousands(next.Value.TrimEnd('.', ','));
                }
                // Số ở dòng TRƯỚC (VD: "380,000\nT.Tiên" — OCR đảo thứ tự)
                if (i - 1 >= 0)
                {
                    var prev = Regex.Match(lines[i - 1].Trim(), @"^[\d][,.\d]*\d[,.]?$");
                    if (prev.Success)
                        return NormalizeToThousands(prev.Value.TrimEnd('.', ','));
                }
            }
            return "";
        }

        /// <summary>
        /// Chuẩn hóa số tiền sang đơn vị nghìn đồng.
        /// VD: "1,200,000" hoặc "1200000" → "1200"
        /// ⚠️ HARDCODED: quy ước 1 đơn vị = 1,000 VND.
        /// </summary>
        private string NormalizeToThousands(string raw)
        {
            // Strip cả dấu phẩy lẫn dấu chấm (cả hai đều được dùng làm dấu ngăn cách nghìn
            // trong hóa đơn VN: "7,280" và "7.280" đều có nghĩa là 7280)
            var digits = raw.Replace(",", "").Replace(".", "");
            if (long.TryParse(digits, out long val))
                return val >= 1000 ? (val / 1000).ToString() : val.ToString();
            return digits;
        }

        private string ExtractDate(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";
            // "tháng" / "thang" / "háng" (OCR bỏ 't') / "hang" / "thang"
            var m1 = Regex.Match(
                text,
                @"ng[aà]y\s+(\d{1,2})\s+(?:th[aá]ng|h[aá]ng)\s+(\d{1,2})\s+n[aă]m\s+(\d{4})",
                RegexOptions.IgnoreCase
            );
            if (m1.Success)
                return $"{m1.Groups[1].Value.PadLeft(2, '0')}-{m1.Groups[2].Value.PadLeft(2, '0')}-{m1.Groups[3].Value}";

            var m2 = Regex.Match(text, @"\b(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})\b");
            if (m2.Success)
                return $"{m2.Groups[1].Value.PadLeft(2, '0')}-{m2.Groups[2].Value.PadLeft(2, '0')}-{m2.Groups[3].Value}";

            return "";
        }

        /// <summary>
        /// Thêm dòng log Gemini vào list (nếu list != null), kèm timestamp.
        /// Caller (OcrTab) sẽ gom list này cùng raw OCR + mapping vào 1 log file thống nhất.
        /// Nếu geminiLog == null thì không ghi gì (backward-compatible).
        /// </summary>
        private static void AddGeminiLog(List<string> geminiLog, string message)
        {
            geminiLog?.Add($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [Gemini] {message}");
        }
    }
}

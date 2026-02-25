using System;
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
        /// <summary>
        /// Extract tất cả 12 fields bắt buộc từ OCR text.
        /// Trả về danh sách các field bị thiếu (empty list = đủ hết).
        /// </summary>
        public List<string> ExtractAllFields(string text, out Dictionary<string, string> fields)
        {
            fields = new Dictionary<string, string>();
            var missingFields = new List<string>();

            // 1. SHOP — lấy dòng bắt đầu "ĐOÀN" nhưng không phải dòng footer (đổi size, ngày kể...)
            // ⚠️ HARDCODED: tên shop cố định theo hóa đơn Đoàn Ngân Châu
            var shopLine = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None)
                               .FirstOrDefault(l =>
                                   (l.Trim().StartsWith("ĐOÀN", StringComparison.OrdinalIgnoreCase)
                                 || l.Trim().StartsWith("Đoàn", StringComparison.OrdinalIgnoreCase))
                                 && !Regex.IsMatch(l, @"nhận đổi|sản phẩm|ngày kể|giặt ủi|khui hàng", RegexOptions.IgnoreCase));
            if (shopLine == null)
            {
                var shopM = Regex.Match(text, @"(?:shop|cửa hàng|store)\s*[:=]?\s*([^\n\r,]+)", RegexOptions.IgnoreCase);
                fields["SHOP"] = shopM.Success ? shopM.Groups[1].Value.Trim() : string.Empty;
            }
            else
            {
                fields["SHOP"] = shopLine.Trim();
            }
            if (string.IsNullOrEmpty(fields["SHOP"])) missingFields.Add("SHOP");

            // 2. TÊN KH — "Khách hàng: ..." / "Khách hàng. ..." / "Khách hàng; ..." (OCR hay đọc ":" thành "."/";")
            var khMatch = Regex.Match(text, @"Kh[aá]ch\s*h[aà]ng\s*[:.=;]?\s*([^\n\r]+)", RegexOptions.IgnoreCase);
            var khRaw = khMatch.Success ? khMatch.Groups[1].Value.Trim() : string.Empty;
            // Strip các ký tự nhiễu đầu: ". ", "; ", "VIP-", "VIP " ...
            khRaw = Regex.Replace(khRaw, @"^[\s;.,\-–—:]+", "").Trim();
            // Nếu có "VIP" ở đầu (như "VIP- TÊN KH") thì bỏ
            khRaw = Regex.Replace(khRaw, @"^VIP[\s\-–—]*", "", RegexOptions.IgnoreCase).Trim();
            fields["TÊN KH"] = khRaw;
            if (string.IsNullOrEmpty(fields["TÊN KH"])) missingFields.Add("TÊN KH");

            // 3. MÃ (Số HĐ) — OCR hay đọc méo "So HD:" thành nhiều dạng khác nhau
            // ⚠️ HARDCODED: format "HD\d+" theo template hóa đơn Đoàn Ngân Châu
            var maField = Regex.Match(text, @"S[oố]\s*H[ĐD\d]\s*[:=]?\s*(HD\d+)", RegexOptions.IgnoreCase);
            if (!maField.Success)
                maField = Regex.Match(text, @"(?:Số\s*HĐ|HĐ\s*số|Invoice\s*No)\s*[:=]?\s*([^\n\r\s,]+)", RegexOptions.IgnoreCase);
            if (!maField.Success)
                // Fallback: tìm "HD" theo sau ngay bởi số — xuất hiện bất kỳ đâu trong dòng
                maField = Regex.Match(text, @"\b(HD\d{4,})\b", RegexOptions.IgnoreCase);
            fields["MÃ"] = maField.Success ? maField.Groups[1].Value.Trim().ToUpper() : string.Empty;
            if (string.IsNullOrEmpty(fields["MÃ"])) missingFields.Add("MÃ");

            // 4–7. Address parts — parsed via AddressParser
            string addressLine = ExtractAddressLine(text);
            var parsed = AddressParser.Parse(addressLine);
            fields["SỐ NHÀ"]    = parsed.SoNha;
            fields["TÊN ĐƯỜNG"] = parsed.TenDuong;
            fields["QUẬN"]      = parsed.Quan;
            // Nếu không parse được SỐ NHÀ + TÊN ĐƯỜNG nhưng có địa chỉ raw → để SỐ NHÀ = raw, TÊN ĐƯỜNG = ""
            // Giúp user thấy có dữ liệu để sửa tay thay vì ô trống hoàn toàn.
            if (string.IsNullOrEmpty(fields["SỐ NHÀ"]) && string.IsNullOrEmpty(fields["TÊN ĐƯỜNG"])
                && !string.IsNullOrEmpty(addressLine))
            {
                fields["SỐ NHÀ"] = addressLine; // raw address — user tự tách
            }
            if (string.IsNullOrEmpty(fields["SỐ NHÀ"]))    missingFields.Add("SỐ NHÀ");
            if (string.IsNullOrEmpty(fields["TÊN ĐƯỜNG"])) missingFields.Add("TÊN ĐƯỜNG");
            if (string.IsNullOrEmpty(fields["QUẬN"]))       missingFields.Add("QUẬN");

            // 8. TIỀN THU — thêm các biến thể OCR đọc méo "tổng thanh toán"
            fields["TIỀN THU"] = ExtractAmountLine(text,
                "tiền thu|thu tiền|tổng thanh toán|tong thanh toan|thanh toán|thanh toan|total");
            if (string.IsNullOrEmpty(fields["TIỀN THU"])) missingFields.Add("TIỀN THU");

            // 9. TIỀN SHIP — optional: nếu không có trong ảnh thì để trống,
            // ProcessImages() sẽ tự tra bảng phí ship theo quận để điền vào.
            fields["TIỀN SHIP"] = ExtractAmountLine(text, "tiền ship|ship|vận chuyển");

            // 10. NGÀY LẤY
            fields["NGÀY LẤY"] = ExtractDate(text);
            if (string.IsNullOrEmpty(fields["NGÀY LẤY"])) missingFields.Add("NGÀY LẤY");

            // NGƯỜI ĐI / NGƯỜI LẤY: được nhập thủ công từ UI, không extract ở đây
            fields["NGƯỜI ĐI"]  = "";
            fields["NGƯỜI LẤY"] = "";

            return missingFields;
        }

        // ─── Private helpers ───────────────────────────────────────────────────

        private string ExtractAddressLine(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
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
                if (line.IndexOf("địa chỉ", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    line.IndexOf("địa chi", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    line.IndexOf("dia chi",  StringComparison.OrdinalIgnoreCase) >= 0 ||
                    line.IndexOf("address",  StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    int colon = line.IndexOf(':');
                    string candidate = colon >= 0 ? line.Substring(colon + 1).Trim() : line.Trim();

                    // Bỏ qua nếu là địa chỉ shop (chứa "CN1", "CN2", "HOTLINE", số điện thoại)
                    if (System.Text.RegularExpressions.Regex.IsMatch(candidate,
                            @"CN\d|HOTLINE|CHUYÊN SỈ|\b09\d{8}\b|\b03\d{8}\b",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                        continue;

                    // Bỏ qua nếu candidate quá ngắn hoặc chỉ là "chi" (OCR đọc nhầm label)
                    if (candidate.Length < 4) continue;

                    found = candidate;
                    foundLineIdx = idx;
                    // Không break — tiếp tục để lấy dòng cuối cùng hợp lệ
                }
            }

            // Nếu dòng "Địa chỉ:" lấy được chỉ toàn ký tự OCR sai (không có chữ số và không có từ địa chỉ),
            // thử ghép với dòng kế tiếp.
            // VD: "Địa chỉ: Foyill tor" (OCR sai) + dòng tiếp "B, 235 nguyễn văn cừ, quận 1"
            // → bỏ phần OCR sai đầu dòng tiếp, lấy từ chữ số đầu tiên trở đi
            if (foundLineIdx >= 0 && !string.IsNullOrEmpty(found))
            {
                bool hasDigit = System.Text.RegularExpressions.Regex.IsMatch(found, @"\d");
                bool hasVietWord = System.Text.RegularExpressions.Regex.IsMatch(found,
                    @"\b(đường|phường|quận|hẻm|ngõ|nguyễn|trần|lê|phú|bình|tân|thành|hồ|hưng|minh|cộng|hòa)\b",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (!hasDigit && !hasVietWord && foundLineIdx + 1 < lines.Length)
                {
                    string nextLine = lines[foundLineIdx + 1].Trim();
                    // Dòng tiếp có thể bắt đầu bằng ký tự OCR sai (VD: "B, 235 nguyễn..."),
                    // tìm vị trí chữ số đầu tiên để bỏ phần rác trước đó
                    var digitMatch = System.Text.RegularExpressions.Regex.Match(nextLine, @"\d");
                    if (digitMatch.Success && nextLine.Length >= 5)
                    {
                        // Lấy từ chữ số đầu tiên (bỏ "B, " hay ký tự OCR sai trước đó)
                        found = nextLine.Substring(digitMatch.Index).Trim();
                        // Strip dấu phẩy/khoảng trắng thừa ở đầu
                        found = found.TrimStart(',', ' ');
                    }
                }
            }

            // Strip trailing garbage: dấu "-", "ạ", "…", khoảng trắng thừa
            found = System.Text.RegularExpressions.Regex.Replace(found, @"[\s\-–—ạ\.…]+$", "").Trim();

            // Strip "TP HCM", "TP. HCM", "Hồ Chí Minh", "Hồ Chí Minh" khỏi cuối địa chỉ
            found = System.Text.RegularExpressions.Regex.Replace(found,
                @",?\s*(?:TP\.?\s*H[CG]M|Hồ\s*Chí\s*Minh|HCM|TP\.?\s*HCM)\s*[ạa]?$", "",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim();

            // Strip prefix "Đc:", "Dc:", "DC:" đầu địa chỉ (VD: "Dc: Số 1 Đinh Lễ...")
            found = System.Text.RegularExpressions.Regex.Replace(found, @"^[Đđ][Cc]\s*:?\s*", "").Trim();

            // Chuẩn hóa "pXqY" / "p.X.qY" thành "pX, qY" để AddressParser split đúng
            // VD: "p10q10" → "p10, q10" | "p7.q5" → "p7, q5"
            found = System.Text.RegularExpressions.Regex.Replace(
                found, @"\b(p\.?\d{1,2})\s*\.?\s*(q\.?\d{1,2})\b", "$1, $2",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            // Strip dấu ":" sau số nhà (VD: "2181: Ng vẫn cư" → "2181 Ng vẫn cư")
            found = System.Text.RegularExpressions.Regex.Replace(found, @"^(\d+)\s*:\s*", "$1 ").Trim();

            // Strip rác sau dấu ". " khi theo sau là chữ hoa (VD: ". Tiệm nail thuỳ thỏ")
            // Giữ lại dấu chấm nếu nó là phần của địa chỉ (VD: "p7.q5") — chỉ strip khi sau ". " có cụm chữ hoa/tên riêng
            found = System.Text.RegularExpressions.Regex.Replace(found, @"\s*\.\s+[A-ZĐÀÁẢÃẠĂẮẶẰẲẴÂẤẦẨẪẬ][^\n]*$", "").Trim();

            // Strip prefix "số" lặp thừa trước địa chỉ (VD: "số 1 Đinh Lễ" — giữ nguyên, không cần)
            return found;
        }

        private string ExtractAmountLine(string text, string keywords)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var keyList = keywords.Split('|');

            for (int i = 0; i < lines.Length; i++)
            {
                foreach (var kw in keyList)
                {
                    if (lines[i].IndexOf(kw, StringComparison.OrdinalIgnoreCase) < 0) continue;
                    // Cho phép dấu chấm/phẩy cuối số (VD: "584,000.")
                    var m = Regex.Match(lines[i], @"[\d][,\d]*\d[,.]?");
                    if (m.Success) return NormalizeToThousands(m.Value.TrimEnd('.', ','));
                    if (i + 1 < lines.Length)
                    {
                        var next = Regex.Match(lines[i + 1].Trim(), @"^[\d][,\d]*\d[,.]?$");
                        if (next.Success) return NormalizeToThousands(next.Value.TrimEnd('.', ','));
                    }
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
            var digits = raw.Replace(",", "");
            if (long.TryParse(digits, out long val))
                return val >= 1000 ? (val / 1000).ToString() : val.ToString();
            return digits;
        }

        private string ExtractDate(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            var m1 = Regex.Match(text,
                @"ng[aà]y\s+(\d{1,2})\s+th[aá]ng\s+(\d{1,2})\s+n[aă]m\s+(\d{4})",
                RegexOptions.IgnoreCase);
            if (m1.Success)
                return $"{m1.Groups[1].Value.PadLeft(2, '0')}-{m1.Groups[2].Value.PadLeft(2, '0')}-{m1.Groups[3].Value}";

            var m2 = Regex.Match(text, @"\b(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})\b");
            if (m2.Success)
                return $"{m2.Groups[1].Value.PadLeft(2, '0')}-{m2.Groups[2].Value.PadLeft(2, '0')}-{m2.Groups[3].Value}";

            return "";
        }
    }
}

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

            // 1. SHOP — trên hóa đơn Đoàn Ngân Châu không có label "shop:", lấy tên sau "ĐOÀN NGÂN CHÂU"
            // ⚠️ HARDCODED: tên shop cố định, hoặc dùng dòng trước "Địa Chi:"
            var shopLine = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None)
                               .FirstOrDefault(l => l.Trim().StartsWith("ĐOÀN", StringComparison.OrdinalIgnoreCase)
                                                 || l.Trim().StartsWith("Đoàn", StringComparison.OrdinalIgnoreCase));
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

            // 2. TÊN KH — "Khách hàng: ..." trên hóa đơn thật
            var khMatch = Regex.Match(text, @"Kh[aá]ch\s*h[aà]ng\s*[:=]?\s*([^\n\r]+)", RegexOptions.IgnoreCase);
            fields["TÊN KH"] = khMatch.Success ? khMatch.Groups[1].Value.Trim() : string.Empty;
            if (string.IsNullOrEmpty(fields["TÊN KH"])) missingFields.Add("TÊN KH");

            // 3. MÃ (Số HĐ) — ưu tiên pattern "So HD: HD\d+" thật sự trên hóa đơn
            // ⚠️ HARDCODED: format "So HD: HD123456" — phụ thuộc template hóa đơn Đoàn Ngân Châu
            var maField = Regex.Match(text, @"So\s*H[ĐD]\s*[:=]?\s*(HD\d+)", RegexOptions.IgnoreCase);
            if (!maField.Success)
                maField = Regex.Match(text, @"(?:Số\s*HĐ|HĐ\s*số|Invoice\s*No)\s*[:=]?\s*([^\n\r\s,]+)", RegexOptions.IgnoreCase);
            fields["MÃ"] = maField.Success ? maField.Groups[1].Value.Trim() : string.Empty;
            if (string.IsNullOrEmpty(fields["MÃ"])) missingFields.Add("MÃ");

            // 4–7. Address parts — parsed via AddressParser
            string addressLine = ExtractAddressLine(text);
            var parsed = AddressParser.Parse(addressLine);
            fields["SỐ NHÀ"]    = parsed.SoNha;
            fields["TÊN ĐƯỜNG"] = parsed.TenDuong;
            fields["QUẬN"]      = parsed.Quan;
            if (string.IsNullOrEmpty(fields["SỐ NHÀ"]))    missingFields.Add("SỐ NHÀ");
            if (string.IsNullOrEmpty(fields["TÊN ĐƯỜNG"])) missingFields.Add("TÊN ĐƯỜNG");
            if (string.IsNullOrEmpty(fields["QUẬN"]))       missingFields.Add("QUẬN");

            // 8. TIỀN THU
            fields["TIỀN THU"] = ExtractAmountLine(text, "tiền thu|thu tiền|tổng thanh toán");
            if (string.IsNullOrEmpty(fields["TIỀN THU"])) missingFields.Add("TIỀN THU");

            // 9. TIỀN SHIP
            fields["TIỀN SHIP"] = ExtractAmountLine(text, "tiền ship|ship|vận chuyển");
            if (string.IsNullOrEmpty(fields["TIỀN SHIP"])) fields["TIỀN SHIP"] = "0";

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
            foreach (var line in lines)
            {
                if (line.IndexOf("địa chỉ", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    line.IndexOf("address",  StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    int colon = line.IndexOf(':');
                    return colon >= 0 ? line.Substring(colon + 1).Trim() : line.Trim();
                }
            }
            return "";
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
                    var m = Regex.Match(lines[i], @"[\d][,\d]*\d");
                    if (m.Success) return NormalizeToThousands(m.Value);
                    if (i + 1 < lines.Length)
                    {
                        var next = Regex.Match(lines[i + 1].Trim(), @"^[\d][,\d]*\d$");
                        if (next.Success) return NormalizeToThousands(next.Value);
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

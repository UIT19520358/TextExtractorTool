using System;
using System.Text.RegularExpressions;

namespace TextInputter.Services
{
    /// <summary>
    /// Service for parsing OCR text and extracting invoice information
    /// </summary>
    public class OCRTextParsingService
    {
        /// <summary>
        /// Extract invoice number from OCR text
        /// </summary>
        public string ExtractInvoiceNumber(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            // TODO: Implement invoice number extraction
            // Common patterns: "HĐ: XXXXX", "Số HĐ: XXXXX", "Invoice: XXXXX"
            var match = Regex.Match(text, @"(?:HĐ|Số HĐ|Invoice|So HD)\s*[:=]?\s*(\d+)", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value : string.Empty;
        }

        /// <summary>
        /// Extract address from OCR text
        /// </summary>
        public string ExtractAddress(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            // TODO: Implement address extraction
            // Look for lines containing address indicators
            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (trimmed.Contains("địa chỉ", StringComparison.OrdinalIgnoreCase) ||
                    trimmed.Contains("address", StringComparison.OrdinalIgnoreCase))
                {
                    return trimmed;
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Extract total amount from OCR text
        /// </summary>
        public decimal ExtractTotalAmount(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0m;

            // TODO: Improve currency amount extraction
            // Pattern for Vietnamese currency: 1.234.567 or 1,234,567
            var matches = Regex.Matches(text, @"(\d+[.,]\d+[.,]?\d*|\d+)");
            
            foreach (Match match in matches)
            {
                var value = match.Value.Replace(".", "").Replace(",", ".");
                if (decimal.TryParse(value, out decimal amount) && amount > 0)
                {
                    return amount;
                }
            }

            return 0m;
        }

        /// <summary>
        /// Extract discount amount from OCR text
        /// </summary>
        public decimal ExtractDiscount(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0m;

            // TODO: Implement discount extraction
            // Look for discount indicators: "giảm", "chiết khấu", "discount", "CK"
            var match = Regex.Match(text, @"(?:giảm|chiết khấu|discount|CK)\s*[:=]?\s*(\d+[.,]\d+|\d+)", RegexOptions.IgnoreCase);
            
            if (match.Success && decimal.TryParse(match.Groups[1].Value.Replace(".", "").Replace(",", "."), out decimal discount))
            {
                return discount;
            }

            return 0m;
        }

        /// <summary>
        /// Extract person name (Người Đi/Người Lấy) from OCR text
        /// </summary>
        public string ExtractPersonName(string text, string personType)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            // TODO: Implement person name extraction
            // Pattern: "Người Đi: Tên" or "Người Lấy: Tên"
            var pattern = personType switch
            {
                "người_đi" => @"(?:người\s*đi|sender)\s*[:=]?\s*([^\n\r]+)",
                "người_lấy" => @"(?:người\s*lấy|receiver)\s*[:=]?\s*([^\n\r]+)",
                _ => null
            };

            if (pattern == null)
                return string.Empty;

            var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value.Trim() : string.Empty;
        }

        /// <summary>
        /// Extract date from OCR text (DD-MM-YYYY or DD/MM/YYYY)
        /// </summary>
        public string ExtractDate(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            var match = Regex.Match(text, @"(\d{1,2})[-/](\d{1,2})[-/](\d{4})");
            return match.Success ? $"{match.Groups[1].Value}-{match.Groups[2].Value}-{match.Groups[3].Value}" : string.Empty;
        }

        /// <summary>
        /// Extract all 12 required fields for OCR Batch Processing
        /// Required: SHOP, TÊN KH, MÃ, SỐ NHÀ, TÊN ĐƯỜNG, QUẬN, TIỀN THU, TIỀN SHIP, TIỀN HÀNG, NGÀY LẤY
        /// Optional: NGƯỜI ĐI, NGƯỜI LẤY (will be input manually in OCR tab)
        /// </summary>
        public List<string> ExtractAllFields(string text, out Dictionary<string, string> fields)
        {
            fields = new Dictionary<string, string>();
            var missingFields = new List<string>();

            // 1. SHOP (extracted from address or invoice header)
            var shop = Regex.Match(text, @"(?:shop|cửa hàng|store)\s*[:=]?\s*([^\n\r,]+)", RegexOptions.IgnoreCase);
            fields["SHOP"] = shop.Success ? shop.Groups[1].Value.Trim() : string.Empty;
            if (string.IsNullOrEmpty(fields["SHOP"])) missingFields.Add("SHOP");

            // 2. TÊN KH (Customer name - similar to address)
            fields["TÊN KH"] = ExtractAddress(text);
            if (string.IsNullOrEmpty(fields["TÊN KH"])) missingFields.Add("TÊN KH");

            // 3. MÃ (Product/Invoice code)
            var maField = Regex.Match(text, @"(?:mã|code|sku)\s*[:=]?\s*([^\n\r\s,]+)", RegexOptions.IgnoreCase);
            fields["MÃ"] = maField.Success ? maField.Groups[1].Value.Trim() : string.Empty;
            if (string.IsNullOrEmpty(fields["MÃ"])) missingFields.Add("MÃ");

            // 4. SỐ NHÀ (House number)
            var soNha = Regex.Match(text, @"(?:số|no\.?|#)\s*[\d]+[a-z]?(?:\s*[/\\-]\s*[\d]+)?", RegexOptions.IgnoreCase);
            fields["SỐ NHÀ"] = soNha.Success ? soNha.Value.Trim() : string.Empty;
            if (string.IsNullOrEmpty(fields["SỐ NHÀ"])) missingFields.Add("SỐ NHÀ");

            // 5. TÊN ĐƯỜNG (Street name)
            var tenDuong = Regex.Match(text, @"(?:đường|street|st\.)\s*([^\n\r,;]+)", RegexOptions.IgnoreCase);
            fields["TÊN ĐƯỜNG"] = tenDuong.Success ? tenDuong.Groups[1].Value.Trim() : string.Empty;
            if (string.IsNullOrEmpty(fields["TÊN ĐƯỜNG"])) missingFields.Add("TÊN ĐƯỜNG");

            // 6. QUẬN (District)
            var quan = Regex.Match(text, @"(?:quận|huyện|q\.)\s*(\d+|[^\n\r,;]+)", RegexOptions.IgnoreCase);
            fields["QUẬN"] = quan.Success ? quan.Groups[1].Value.Trim() : string.Empty;
            if (string.IsNullOrEmpty(fields["QUẬN"])) missingFields.Add("QUẬN");

            // 7. TIỀN THU (Amount collected - use ExtractTotalAmount)
            var tienThu = ExtractTotalAmount(text);
            fields["TIỀN THU"] = tienThu > 0 ? tienThu.ToString() : string.Empty;
            if (tienThu <= 0) missingFields.Add("TIỀN THU");

            // 8. TIỀN SHIP (Shipping cost)
            var tienShip = Regex.Match(text, @"(?:phí\s*ship|ship|shipping)\s*[:=]?\s*([\d,\.]+)", RegexOptions.IgnoreCase);
            if (tienShip.Success && decimal.TryParse(tienShip.Groups[1].Value.Replace(",", ""), out var shipAmount))
            {
                fields["TIỀN SHIP"] = shipAmount > 0 ? shipAmount.ToString() : string.Empty;
            }
            else
            {
                fields["TIỀN SHIP"] = string.Empty;
            }
            if (string.IsNullOrEmpty(fields["TIỀN SHIP"])) missingFields.Add("TIỀN SHIP");

            // 9. TIỀN HÀNG (Product cost)
            var tienHang = Regex.Match(text, @"(?:tiền\s*hàng|product|cost)\s*[:=]?\s*([\d,\.]+)", RegexOptions.IgnoreCase);
            if (tienHang.Success && decimal.TryParse(tienHang.Groups[1].Value.Replace(",", ""), out var productAmount))
            {
                fields["TIỀN HÀNG"] = productAmount > 0 ? productAmount.ToString() : string.Empty;
            }
            else
            {
                fields["TIỀN HÀNG"] = string.Empty;
            }
            if (string.IsNullOrEmpty(fields["TIỀN HÀNG"])) missingFields.Add("TIỀN HÀNG");

            // 10. NGÀY LẤY (Pickup date)
            fields["NGÀY LẤY"] = ExtractDate(text);
            if (string.IsNullOrEmpty(fields["NGÀY LẤY"])) missingFields.Add("NGÀY LẤY");

            // 11. NGƯỜI ĐI (Sender) - OPTIONAL - will be input manually in OCR tab
            fields["NGƯỜI ĐI"] = ExtractPersonName(text, "người_đi");
            // Not in missing list - will be handled manually

            // 12. NGƯỜI LẤY (Receiver) - OPTIONAL - will be input manually in OCR tab
            fields["NGƯỜI LẤY"] = ExtractPersonName(text, "người_lấy");
            // Not in missing list - will be handled manually

            return missingFields;
        }
    }
}

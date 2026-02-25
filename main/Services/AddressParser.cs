using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace TextInputter.Services
{
    /// <summary>
    /// Lớp hỗ trợ parsing địa chỉ Việt Nam
    /// Format chuẩn: Số nhà | Tên đường | Phường | Quận
    /// </summary>
    public class AddressParser
    {
        // Dictionary các quận/huyện phổ biến
        private static readonly Dictionary<string, string> DistrictDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // TP HCM - số thuần
            { "1", "1" }, { "2", "2" }, { "3", "3" }, { "4", "4" },
            { "5", "5" }, { "6", "6" }, { "7", "7" }, { "8", "8" },
            { "9", "9" }, { "10", "10" }, { "11", "11" }, { "12", "12" },
            // Prefix q/q.
            { "q1", "1" }, { "q.1", "1" }, { "quận 1", "1" }, { "quan 1", "1" },
            { "q2", "2" }, { "q.2", "2" }, { "quận 2", "2" }, { "quan 2", "2" },
            { "q3", "3" }, { "q.3", "3" }, { "quận 3", "3" }, { "quan 3", "3" },
            { "q4", "4" }, { "q.4", "4" }, { "quận 4", "4" }, { "quan 4", "4" },
            { "q5", "5" }, { "q.5", "5" }, { "quận 5", "5" }, { "quan 5", "5" },
            { "q6", "6" }, { "q.6", "6" }, { "quận 6", "6" }, { "quan 6", "6" },
            { "q7", "7" }, { "q.7", "7" }, { "quận 7", "7" }, { "quan 7", "7" },
            { "q8", "8" }, { "q.8", "8" }, { "quận 8", "8" }, { "quan 8", "8" },
            { "q9", "9" }, { "q.9", "9" }, { "quận 9", "9" }, { "quan 9", "9" },
            { "q10", "10" }, { "q.10", "10" }, { "quận 10", "10" }, { "quan 10", "10" },
            { "q11", "11" }, { "q.11", "11" }, { "quận 11", "11" }, { "quan 11", "11" },
            { "q12", "12" }, { "q.12", "12" }, { "quận 12", "12" }, { "quan 12", "12" },
            // Phú Nhuận — bao gồm các biến thể OCR sai dấu phổ biến
            { "phú nhuận", "phu nhuan" }, { "phủ nhuận", "phu nhuan" },
            { "phú nhuật", "phu nhuan" }, { "phủ nhuật", "phu nhuan" },
            { "phu nhuat", "phu nhuan" }, { "phu nhuai", "phu nhuan" },
            { "phú nhuăn", "phu nhuan" }, { "phu nhuan", "phu nhuan" },
            // Bình Thạnh
            { "bình thạnh", "binh thanh" }, { "bình thanh", "binh thanh" },
            { "binh thạnh", "binh thanh" }, { "binh thanh", "binh thanh" },
            // Bình Tân
            { "bình tân", "binh tan" }, { "binh tân", "binh tan" },
            { "bình tan", "binh tan" }, { "binh tan", "binh tan" },
            // Tân Phú
            { "tân phú", "tan phu" }, { "tan phú", "tan phu" },
            { "tân phu", "tan phu" }, { "tan phu", "tan phu" },
            // Tân Bình
            { "tân bình", "tan binh" }, { "tân binh", "tan binh" },
            { "tan bình", "tan binh" }, { "tan binh", "tan binh" },
            // Gò Vấp
            { "gò vấp", "go vap" }, { "go vấp", "go vap" },
            { "gò vap", "go vap" }, { "go vap", "go vap" },
            // Thủ Đức
            { "thủ đức", "thu duc" }, { "thu đức", "thu duc" },
            { "thủ duc", "thu duc" }, { "thu duc", "thu duc" },
            // Bình Chánh
            { "bình chánh", "binh chanh" }, { "binh chánh", "binh chanh" },
            { "bình chanh", "binh chanh" }, { "binh chanh", "binh chanh" },
            // Hóc Môn
            { "hóc môn", "hoc mon" }, { "hoc môn", "hoc mon" },
            { "hóc mon", "hoc mon" }, { "hoc mon", "hoc mon" },
            // Củ Chi
            { "củ chi", "cu chi" }, { "cu chi", "cu chi" },
            // Nhà Bè
            { "nhà bè", "nha be" }, { "nha be", "nha be" },
            // Cần Giờ
            { "cần giờ", "can gio" }, { "can gio", "can gio" },
            // Hà Nội
            { "tây hồ", "tay ho" },
            { "hoàn kiếm", "hoan kiem" },
            { "ba đình", "ba dinh" },
            { "đống đa", "dong da" },
            { "hai bà trưng", "hai ba trung" },
            { "thanh xuân", "thanh xuan" },
            { "cầu giấy", "cau giay" },
            // Tỉnh
            { "tinh", "tinh" }, { "tỉnh", "tinh" },
        };

        // Dictionary phường
        private static readonly Dictionary<string, string> WardDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "p1", "Phường 1" }, { "p.1", "Phường 1" }, { "phường 1", "Phường 1" }, { "phuong 1", "Phường 1" },
            { "p2", "Phường 2" }, { "p.2", "Phường 2" }, { "phường 2", "Phường 2" }, { "phuong 2", "Phường 2" },
            { "p3", "Phường 3" }, { "p.3", "Phường 3" }, { "phường 3", "Phường 3" }, { "phuong 3", "Phường 3" },
            { "p4", "Phường 4" }, { "p.4", "Phường 4" }, { "phường 4", "Phường 4" }, { "phuong 4", "Phường 4" },
            { "p5", "Phường 5" }, { "p.5", "Phường 5" }, { "phường 5", "Phường 5" }, { "phuong 5", "Phường 5" },
            { "p6", "Phường 6" }, { "p.6", "Phường 6" }, { "phường 6", "Phường 6" }, { "phuong 6", "Phường 6" },
            { "p7", "Phường 7" }, { "p.7", "Phường 7" }, { "phường 7", "Phường 7" }, { "phuong 7", "Phường 7" },
            { "p8", "Phường 8" }, { "p.8", "Phường 8" }, { "phường 8", "Phường 8" }, { "phuong 8", "Phường 8" },
            { "p9", "Phường 9" }, { "p.9", "Phường 9" }, { "phường 9", "Phường 9" }, { "phuong 9", "Phường 9" },
            { "p10", "Phường 10" }, { "p.10", "Phường 10" }, { "phường 10", "Phường 10" }, { "phuong 10", "Phường 10" },
            { "p11", "Phường 11" }, { "p.11", "Phường 11" }, { "phường 11", "Phường 11" }, { "phuong 11", "Phường 11" },
            { "p12", "Phường 12" }, { "p.12", "Phường 12" }, { "phường 12", "Phường 12" }, { "phuong 12", "Phường 12" },
            { "p13", "Phường 13" }, { "p.13", "Phường 13" }, { "phường 13", "Phường 13" }, { "phuong 13", "Phường 13" },
            { "p14", "Phường 14" }, { "p.14", "Phường 14" }, { "phường 14", "Phường 14" }, { "phuong 14", "Phường 14" },
            { "p15", "Phường 15" }, { "p.15", "Phường 15" }, { "phường 15", "Phường 15" }, { "phuong 15", "Phường 15" },
        };

        /// <summary>
        /// Cấu trúc kết quả parse địa chỉ
        /// </summary>
        public class ParsedAddress
        {
            public string SoNha { get; set; } = "";
            public string TenDuong { get; set; } = "";
            public string Phuong { get; set; } = "";
            public string Quan { get; set; } = "";
            public float Confidence { get; set; } = 0f; // 0-1 (0 = không chắc chắn)
        }

        /// <summary>
        /// Parse địa chỉ từ string
        /// Trả về ParsedAddress với confidence score
        /// </summary>
        public static ParsedAddress Parse(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
                return new ParsedAddress { Confidence = 0 };

            var result = new ParsedAddress();
            address = address.Trim();

            // Chuẩn hóa: chèn dấu phẩy trước các token phường/quận inline (không có dấu phẩy)
            // VD: "5/1 phùng văn cung p2 phủ nhuận"
            //   → "5/1 phùng văn cung, p2, phủ nhuận"
            // VD: "133/3 hoà hưng p12 q10"
            //   → "133/3 hoà hưng, p12, q10"
            // VD: "458/59 lý thái tổ p10q10"  (không có dấu cách)
            //   → "458/59 lý thái tổ, p10, q10"
            // Bước 1: tách p10q10 → p10, q10
            address = Regex.Replace(address, @"(?<=[a-zA-ZÀ-ỹ0-9])(p\.?\d{1,2})(q\.?\d{1,2})\b", ", $1, $2", RegexOptions.IgnoreCase);
            // Bước 2: chèn phẩy trước "p<số>" khi đứng sau chữ/số (không phải đầu chuỗi hoặc sau phẩy)
            address = Regex.Replace(address, @"(?<=\S)\s+(p\.?\s*\d{1,2})\b(?!\s*\d)", ", $1", RegexOptions.IgnoreCase);
            // Bước 3: chèn phẩy trước "q<số>" khi đứng sau chữ/số
            address = Regex.Replace(address, @"(?<=\S)\s+(q\.?\s*\d{1,2})\b", ", $1", RegexOptions.IgnoreCase);
            // Bước 4: chèn phẩy trước "phường <tên>" khi đứng sau chữ/số (bắt cả tên phường nhiều từ)
            address = Regex.Replace(address, @"(?<=\S)\s+(ph[uướừửữ][oôờ]ng\s+[^\s,]+(?:\s+[^\s,]+)*)", ", $1", RegexOptions.IgnoreCase);
            // Bước 5: chèn phẩy trước "quận <tên/số>" khi đứng sau chữ/số
            address = Regex.Replace(address, @"(?<=\S)\s+(qu[aâậ]n\s+[^\s,]+(?:\s+[^\s,]+)*)", ", $1", RegexOptions.IgnoreCase);

            // Split bởi dấu phẩy để xử lý từng segment
            // VD: "A25 hotel (phòng 706) 184 nguyễn trãi, phường phạm ngũ lão, q1"
            //   → segments = ["A25 hotel (phòng 706) 184 nguyễn trãi", "phường phạm ngũ lão", "q1"]
            var segments = address.Split(',')
                                  .Select(s => s.Trim())
                                  .Where(s => !string.IsNullOrEmpty(s))
                                  .ToList();

            // Bước 1: Tìm quận ở các segment (ưu tiên từ cuối)
            int districtSegIdx = -1;
            for (int i = segments.Count - 1; i >= 0; i--)
            {
                var (quan, _) = FindDistrictInSegment(segments[i]);
                if (!string.IsNullOrEmpty(quan))
                {
                    result.Quan = quan;
                    result.Confidence += 0.3f;
                    districtSegIdx = i;
                    break;
                }
            }
            if (districtSegIdx >= 0) segments.RemoveAt(districtSegIdx);

            // Bước 2: Tìm phường ở các segment còn lại (ưu tiên từ cuối)
            int wardSegIdx = -1;
            for (int i = segments.Count - 1; i >= 0; i--)
            {
                var (phuong, _) = FindWardInSegment(segments[i]);
                if (!string.IsNullOrEmpty(phuong))
                {
                    result.Phuong = phuong;
                    result.Confidence += 0.3f;
                    wardSegIdx = i;
                    break;
                }
            }
            if (wardSegIdx >= 0) segments.RemoveAt(wardSegIdx);

            // Bước 3: Segment đầu tiên còn lại = "số nhà + tên đường"
            string remaining = string.Join(", ", segments).Trim();
            var (soNha, tenDuong) = ExtractHouseAndStreet(remaining);
            result.SoNha = soNha;
            result.TenDuong = CleanText(tenDuong);
            if (!string.IsNullOrEmpty(result.SoNha)) result.Confidence += 0.2f;
            if (!string.IsNullOrEmpty(result.TenDuong)) result.Confidence += 0.2f;

            result.Confidence = Math.Min(result.Confidence, 1f);

            // Chuẩn hóa output về lowercase không dấu để đồng nhất với data nhập tay.
            // Quận số ("1".."12") giữ nguyên; quận tên → lowercase ("Binh Thanh" → "binh thanh").
            result.Quan     = NormalizeOutput(result.Quan);
            result.TenDuong = NormalizeOutput(result.TenDuong);
            result.SoNha    = NormalizeOutput(result.SoNha);
            result.Phuong   = NormalizeOutput(result.Phuong);

            return result;
        }

        /// <summary>
        /// Chuẩn hóa output: lowercase, không dấu, trim.
        /// Số quận ("1".."12") giữ nguyên vì đã không dấu.
        /// </summary>
        private static string NormalizeOutput(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return s;
            return RemoveDiacritics(s.Trim()).ToLowerInvariant();
        }

        /// <summary>
        /// Xóa dấu tiếng Việt (NFD decompose + loại combining marks).
        /// </summary>
        private static string RemoveDiacritics(string s)
        {
            var normalized = s.Normalize(System.Text.NormalizationForm.FormD);
            var sb = new System.Text.StringBuilder();
            foreach (char c in normalized)
            {
                if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c)
                    != System.Globalization.UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            }
            return sb.ToString().Normalize(System.Text.NormalizationForm.FormC);
        }

        /// <summary>
        /// Tìm quận/huyện trong text
        /// Trả về (Tên quận, Text gốc tìm được)
        /// </summary>
        private static (string, string) FindDistrict(string text)
        {
            return FindDistrictInSegment(text);
        }

        // Lazy-initialized: map key không dấu → giá trị, dùng cho fuzzy lookup OCR sai dấu
        private static Dictionary<string, string> _districtNoDiacDict;
        private static Dictionary<string, string> DistrictNoDiacDict
        {
            get
            {
                if (_districtNoDiacDict != null) return _districtNoDiacDict;
                _districtNoDiacDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var kv in DistrictDict)
                {
                    var keyNoDiac = RemoveDiacriticsStatic(kv.Key);
                    if (!_districtNoDiacDict.ContainsKey(keyNoDiac))
                        _districtNoDiacDict[keyNoDiac] = kv.Value;
                }
                return _districtNoDiacDict;
            }
        }

        /// <summary>
        /// Tìm quận trong 1 segment (không có dấu phẩy)
        /// </summary>
        private static (string, string) FindDistrictInSegment(string seg)
        {
            seg = seg.Trim();
            if (string.IsNullOrEmpty(seg)) return ("", "");

            // Match: q1, q.1, quận 1, quan 1 — cả segment hoặc cuối segment
            var mQ = Regex.Match(seg, @"\bqu[aâ]n\.?\s*(\d{1,2})\b", RegexOptions.IgnoreCase);
            if (!mQ.Success)
                mQ = Regex.Match(seg, @"\bq\.?\s*(\d{1,2})\b", RegexOptions.IgnoreCase);
            if (mQ.Success)
                return (mQ.Groups[1].Value, mQ.Value.Trim());

            // Lookup toàn bộ segment (có dấu)
            if (DistrictDict.TryGetValue(seg, out var d1))
                return (d1, seg);

            // Lookup từng word + cặp 2 từ (có dấu)
            var words = seg.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = words.Length - 1; i >= 0; i--)
            {
                if (DistrictDict.TryGetValue(words[i], out var d2))
                    return (d2, words[i]);
                if (i > 0)
                {
                    var combined = words[i - 1] + " " + words[i];
                    if (DistrictDict.TryGetValue(combined, out var d3))
                        return (d3, combined);
                }
                if (i > 1)
                {
                    var triple = words[i - 2] + " " + words[i - 1] + " " + words[i];
                    if (DistrictDict.TryGetValue(triple, out var d4))
                        return (d4, triple);
                }
            }

            // Fuzzy: xóa dấu → lookup trong DistrictNoDiacDict
            // Bắt OCR đọc sai dấu: "phủ nhuận"→"phu nhuan", "phú nhuật"→"phu nhuat" đều → "phu nhuan"
            var segNoDiac = RemoveDiacriticsStatic(seg);
            if (DistrictNoDiacDict.TryGetValue(segNoDiac, out var df1))
                return (df1, seg);

            var wordsNoDiac = segNoDiac.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = wordsNoDiac.Length - 1; i >= 0; i--)
            {
                if (DistrictNoDiacDict.TryGetValue(wordsNoDiac[i], out var df2))
                    return (df2, wordsNoDiac[i]);
                if (i > 0)
                {
                    var combined = wordsNoDiac[i - 1] + " " + wordsNoDiac[i];
                    if (DistrictNoDiacDict.TryGetValue(combined, out var df3))
                        return (df3, combined);
                }
                if (i > 1)
                {
                    var triple = wordsNoDiac[i - 2] + " " + wordsNoDiac[i - 1] + " " + wordsNoDiac[i];
                    if (DistrictNoDiacDict.TryGetValue(triple, out var df4))
                        return (df4, triple);
                }
            }

            return ("", "");
        }

        /// <summary>
        /// Xóa dấu tiếng Việt — dùng nội bộ trong AddressParser để fuzzy match.
        /// </summary>
        private static string RemoveDiacriticsStatic(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            var normalized = s.Normalize(System.Text.NormalizationForm.FormD);
            var sb = new System.Text.StringBuilder();
            foreach (char c in normalized)
                if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c)
                    != System.Globalization.UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            return sb.ToString().Normalize(System.Text.NormalizationForm.FormC).ToLowerInvariant();
        }

        /// <summary>
        /// Tìm phường/tổ trong text
        /// </summary>
        private static (string, string) FindWard(string text)
        {
            return FindWardInSegment(text);
        }

        /// <summary>
        /// Tìm phường trong 1 segment — nhận cả p1/p.1/phường 1 lẫn tên phường dài
        /// </summary>
        private static (string, string) FindWardInSegment(string seg)
        {
            seg = seg.Trim();
            if (string.IsNullOrEmpty(seg)) return ("", "");

            // Lookup trong WardDict trước (exact match & contains)
            foreach (var key in WardDict.Keys)
            {
                if (seg.Equals(key, StringComparison.OrdinalIgnoreCase))
                    return (WardDict[key], seg);
            }
            foreach (var key in WardDict.Keys)
            {
                if (seg.Contains(key, StringComparison.OrdinalIgnoreCase))
                    return (WardDict[key], key);
            }

            // Regex fallback: bắt "p12", "p.12", "phường 12", "phuong 12" — bất kỳ số nào
            var mPNum = Regex.Match(seg, @"\b(?:ph?u[oô]ng\.?\s*|p\.?\s*)(\d{1,2})\b", RegexOptions.IgnoreCase);
            if (mPNum.Success)
                return ($"Phường {mPNum.Groups[1].Value}", mPNum.Value.Trim());

            // Nhận dạng "phường <tên>" hoặc "p. <tên>" dù tên không có trong dict
            var lower = seg.ToLower();
            if (lower.StartsWith("phường") || lower.StartsWith("phuong") || lower.StartsWith("p."))
                return (seg, seg);

            return ("", "");
        }

        /// <summary>
        /// Extract số nhà và tên đường từ phần địa chỉ còn lại (sau khi đã tách phường/quận)
        /// Logic: Split theo dấu phẩy → segment đầu tiên chứa số = Số nhà + tên đường trong cùng segment đó
        /// VD: "A25 hotel ( phòng 706) 184 nguyễn trãi" → SoNha="A25 hotel ( phòng 706) 184", TenDuong="nguyễn trãi"
        /// VD: "132 bên Vân đồn" → SoNha="132", TenDuong="bên Vân đồn"
        /// </summary>
        private static (string soNha, string tenDuong) ExtractHouseAndStreet(string address)
        {
            // address ở đây đã bỏ phường, quận
            // Lấy segment đầu (trước dấu phẩy đầu tiên nếu có)
            var firstSeg = address.Split(',')[0].Trim();

            // Ưu tiên: nhận dạng "số <N> đường <tên>" hoặc "<N> đường <tên>"
            // VD: "số 28 đường số 4 khu 2756"  → SoNha="28", TenDuong="đường số 4"
            // VD: "28 đường số 4"               → SoNha="28", TenDuong="đường số 4"
            var duongMatch = Regex.Match(firstSeg,
                @"^(?:số\s*)?(\d+(?:/\d+)?[A-Z]*)\s+(đường\s+.+?)(?:\s+khu\b|\s+lô\b|\s+kdc\b|\s*$)",
                RegexOptions.IgnoreCase);
            if (duongMatch.Success)
            {
                return (duongMatch.Groups[1].Value.Trim(), duongMatch.Groups[2].Value.Trim());
            }

            // Tìm số (dãy số cuối cùng trong segment đầu) — đây là số nhà "thực"
            // Pattern: phần từ đầu đến sau số cuối là SoNha, phần còn lại là TenDuong
            // VD: "A25 hotel ( phòng 706) 184 nguyễn trãi"
            //      → Group1 = "A25 hotel ( phòng 706) 184", Group2 = "nguyễn trãi"
            var numMatch = Regex.Match(firstSeg, @"^(.*\d+[A-Z]?)\s+([^\d(].*)$", RegexOptions.IgnoreCase);
            if (numMatch.Success)
            {
                var soNha = numMatch.Groups[1].Value.Trim();
                var tenDuong = numMatch.Groups[2].Value.Trim();
                return (soNha, tenDuong);
            }

            // Fallback: regex cũ — chỉ lấy số đầu
            var simpleMatch = Regex.Match(firstSeg, @"^(?:số\s*)?(\d+(?:/\d+)?[A-Z]*)\s*(.*)", RegexOptions.IgnoreCase);
            if (simpleMatch.Success)
                return (simpleMatch.Groups[1].Value.Trim(), simpleMatch.Groups[2].Value.Trim());

            return ("", firstSeg);
        }

        /// <summary>
        /// Extract số nhà (thường bắt đầu bằng digit hoặc "số")
        /// VD: "5/1", "123", "số 45A", etc
        /// </summary>
        private static string ExtractHouseNumber(string address)
        {
            // Regex: digit + optional "/" + digit + optional letters
            var match = Regex.Match(address, @"^(?:số\s*)?(\d+(?:/\d+)?[A-Z]*)", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;

            return "";
        }

        /// <summary>
        /// Clean text: trim, normalize spaces
        /// </summary>
        private static string CleanText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            // Xóa multiple spaces, trim
            text = Regex.Replace(text, @"\s+", " ").Trim();
            // Xóa dấu phẩy, dấu chấm ở cuối
            text = text.TrimEnd(',', '.', ' ');

            return text;
        }

        /// <summary>
        /// Format địa chỉ lại theo chuẩn
        /// </summary>
        public static string FormatAddress(ParsedAddress addr)
        {
            var parts = new List<string>();
            if (!string.IsNullOrEmpty(addr.SoNha)) parts.Add(addr.SoNha);
            if (!string.IsNullOrEmpty(addr.TenDuong)) parts.Add(addr.TenDuong);
            if (!string.IsNullOrEmpty(addr.Phuong)) parts.Add(addr.Phuong);
            if (!string.IsNullOrEmpty(addr.Quan)) parts.Add(addr.Quan);

            return string.Join(", ", parts.Where(p => !string.IsNullOrEmpty(p)));
        }
    }
}

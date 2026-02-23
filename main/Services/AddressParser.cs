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
            // Tên quận
            { "phú nhuận", "Phú Nhuận" }, { "phu nhuan", "Phú Nhuận" },
            { "bình thạnh", "Bình Thạnh" }, { "binh thanh", "Bình Thạnh" },
            { "bình tân", "Bình Tân" }, { "binh tan", "Bình Tân" },
            { "tân phú", "Tân Phú" }, { "tan phu", "Tân Phú" },
            { "tân bình", "Tân Bình" }, { "tan binh", "Tân Bình" },
            { "gò vấp", "Gò Vấp" }, { "go vap", "Gò Vấp" },
            { "phú xuân", "Phú Xuân" },
            { "thủ đức", "Thủ Đức" }, { "thu duc", "Thủ Đức" },
            { "tây hồ", "Tây Hồ" },
            { "hoàn kiếm", "Hoàn Kiếm" },
            { "ba đình", "Ba Đình" },
            { "đống đa", "Đống Đa" },
            { "hai bà trưng", "Hai Bà Trưng" },
            { "thanh xuân", "Thanh Xuân" },
            { "cầu giấy", "Cầu Giấy" },
            // Tỉnh
            { "tinh", "tinh" }, { "tỉnh", "tinh" },
        };

        // Dictionary phường
        private static readonly Dictionary<string, string> WardDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "p1", "Phường 1" }, { "p.1", "Phường 1" }, { "phường 1", "Phường 1" },
            { "p2", "Phường 2" }, { "p.2", "Phường 2" }, { "phường 2", "Phường 2" },
            { "p3", "Phường 3" }, { "p.3", "Phường 3" }, { "phường 3", "Phường 3" },
            { "p4", "Phường 4" }, { "p.4", "Phường 4" }, { "phường 4", "Phường 4" },
            { "p5", "Phường 5" }, { "p.5", "Phường 5" }, { "phường 5", "Phường 5" },
            { "p6", "Phường 6" }, { "p.6", "Phường 6" }, { "phường 6", "Phường 6" },
            { "p7", "Phường 7" }, { "p.7", "Phường 7" }, { "phường 7", "Phường 7" },
            { "p8", "Phường 8" }, { "p.8", "Phường 8" }, { "phường 8", "Phường 8" },
            { "p9", "Phường 9" }, { "p.9", "Phường 9" }, { "phường 9", "Phường 9" },
            { "p10", "Phường 10" }, { "p.10", "Phường 10" }, { "phường 10", "Phường 10" },
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
            return result;
        }

        /// <summary>
        /// Tìm quận/huyện trong text
        /// Trả về (Tên quận, Text gốc tìm được)
        /// </summary>
        private static (string, string) FindDistrict(string text)
        {
            return FindDistrictInSegment(text);
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

            // Lookup toàn bộ segment
            if (DistrictDict.TryGetValue(seg, out var d1))
                return (d1, seg);

            // Lookup từng word + cặp 2 từ (cho "phú nhuận", "bình thạnh"...)
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
            }

            return ("", "");
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

            // Lookup trong WardDict trước
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

            // Nhận dạng "phường <tên>" hoặc "p. <tên>" dù tên không có trong dict
            var mPhuong = Regex.Match(seg, @"^ph?\.?\s*(.+)$", RegexOptions.IgnoreCase);
            if (mPhuong.Success && seg.Length > 2)
            {
                // Chỉ nhận nếu bắt đầu bằng "ph" hoặc "p." (không phải phần của địa chỉ chính)
                var lower = seg.ToLower();
                if (lower.StartsWith("phường") || lower.StartsWith("phuong") ||
                    lower.StartsWith("p.") || Regex.IsMatch(seg, @"^p\d", RegexOptions.IgnoreCase))
                {
                    return (seg, seg); // Giữ nguyên tên phường
                }
            }

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

            // Tìm số (dãy số cuối cùng trong segment đầu) — đây là số nhà "thực"
            // Pattern: phần từ đầu đến sau số cuối là SoNha, phần còn lại là TenDuong
            // VD: "A25 hotel ( phòng 706) 184 nguyễn trãi"
            //      → số cuối: "184" tại index X
            //      → SoNha = "A25 hotel ( phòng 706) 184"
            //      → TenDuong = "nguyễn trãi"

            // Tìm vị trí kết thúc của số nhà cuối cùng (pattern: digits có thể theo sau là ký tự A-Z)
            // Dùng greedy .* để lấy đến số CUỐI CÙNG, phần sau = tên đường (bắt đầu bằng chữ)
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

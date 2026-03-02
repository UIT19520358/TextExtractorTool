using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using TextInputter;

namespace TextInputter.Services
{
    public class AddressParser
    {
        private static readonly Dictionary<string, string> DistrictDict = new Dictionary<
            string,
            string
        >(StringComparer.OrdinalIgnoreCase)
        {
            { "1", "1" },
            { "2", "2" },
            { "3", "3" },
            { "4", "4" },
            { "5", "5" },
            { "6", "6" },
            { "7", "7" },
            { "8", "8" },
            { "9", "9" },
            { "10", "10" },
            { "11", "11" },
            { "12", "12" },
            { "phu nhuan", "phu nhuan" },
            { "binh thanh", "binh thanh" },
            { "binh tan", "binh tan" },
            { "tan phu", "tan phu" },
            { "tan binh", "tan binh" },
            { "go vap", "go vap" },
            { "thu duc", "thu duc" },
            { "binh chanh", "binh chanh" },
            { "hoc mon", "hoc mon" },
            { "cu chi", "cu chi" },
            { "nha be", "nha be" },
            { "can gio", "can gio" },
            { "tay ho", "tay ho" },
            { "hoan kiem", "hoan kiem" },
            { "ba dinh", "ba dinh" },
            { "dong da", "dong da" },
            { "hai ba trung", "hai ba trung" },
            { "thanh xuan", "thanh xuan" },
            { "cau giay", "cau giay" },
            { "tinh", "tinh" },
        };

        private static readonly Dictionary<string, string> DistrictAliasDict = new Dictionary<
            string,
            string
        >(StringComparer.OrdinalIgnoreCase)
        {
            { "bthanh", "binh thanh" },
            { "tphu", "tan phu" },
            { "tbinh", "tan binh" },
            { "gvap", "go vap" },
            { "tduc", "thu duc" },
            { "pnhuan", "phu nhuan" },
            { "btan", "binh tan" },
        };

        private static readonly Dictionary<string, string> WardNameDict = new Dictionary<
            string,
            string
        >(StringComparer.OrdinalIgnoreCase)
        { };

        private static string NormalizeKey(string s)
        {
            if (string.IsNullOrEmpty(s))
                return "";
            var noDiac = RemoveDiacriticsStatic(s);
            return Regex.Replace(noDiac, @"[^a-z0-9]", "");
        }

        private static Dictionary<string, string> _districtNormalizedDict;
        private static Dictionary<string, string> DistrictNormalizedDict
        {
            get
            {
                if (_districtNormalizedDict != null)
                    return _districtNormalizedDict;
                _districtNormalizedDict = new Dictionary<string, string>(
                    StringComparer.OrdinalIgnoreCase
                );
                foreach (var kv in DistrictDict)
                {
                    var normKey = NormalizeKey(kv.Key);
                    if (!_districtNormalizedDict.ContainsKey(normKey))
                        _districtNormalizedDict[normKey] = kv.Value;
                }
                foreach (var kv in DistrictAliasDict)
                {
                    var normKey = NormalizeKey(kv.Key);
                    if (!_districtNormalizedDict.ContainsKey(normKey))
                        _districtNormalizedDict[normKey] = kv.Value;
                }
                return _districtNormalizedDict;
            }
        }

        // Ward → District map đã normalize key (xóa dấu + space) để tra nhanh
        private static Dictionary<string, string> _wardNormalizedDict;
        private static Dictionary<string, string> WardNormalizedDict
        {
            get
            {
                if (_wardNormalizedDict != null)
                    return _wardNormalizedDict;
                _wardNormalizedDict = new Dictionary<string, string>(
                    StringComparer.OrdinalIgnoreCase
                );
                foreach (var kv in AppConstants.WARD_TO_DISTRICT_MAP)
                {
                    var normKey = NormalizeKey(kv.Key);
                    if (!_wardNormalizedDict.ContainsKey(normKey))
                        _wardNormalizedDict[normKey] = kv.Value;
                }
                return _wardNormalizedDict;
            }
        }

        private static List<string> _knownDistrictEndPatterns;

        private static List<string> BuildKnownDistrictEndPatterns()
        {
            if (_knownDistrictEndPatterns != null)
                return _knownDistrictEndPatterns;
            _knownDistrictEndPatterns = new List<string>();

            // Tên quận có dấu (để match địa chỉ gõ/OCR đầy đủ dấu) — ≥2 từ
            var withDiac = new[]
            {
                "tân phú",
                "tân bình",
                "bình thạnh",
                "bình tân",
                "gò vấp",
                "thủ đức",
                "bình chánh",
                "hóc môn",
                "củ chi",
                "nhà bè",
                "cần giờ",
                "phú nhuận",
                "tây hồ",
                "hoàn kiếm",
                "ba đình",
                "đống đa",
                "hai bà trưng",
                "thanh xuân",
                "cầu giấy",
            };
            // Tên quận không dấu (từ DistrictDict key) — ≥2 từ
            var noDiac = DistrictDict
                .Keys.Where(k => !Regex.IsMatch(k, @"^\d{1,2}$"))
                .Where(k => k.Contains(' '))
                .ToList();

            foreach (var d in withDiac.Concat(noDiac).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var pattern = @"(?<=[^\s,])\s+(" + Regex.Escape(d) + @")\s*$";
                _knownDistrictEndPatterns.Add(pattern);
            }
            return _knownDistrictEndPatterns;
        }

        public class ParsedAddress
        {
            public string SoNha { get; set; } = "";
            public string TenDuong { get; set; } = "";
            public string Phuong { get; set; } = "";
            public string Quan { get; set; } = "";
            public float Confidence { get; set; } = 0f;
        }

        public static ParsedAddress Parse(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
                return new ParsedAddress { Confidence = 0 };

            var result = new ParsedAddress();
            address = address.Trim();

            address = Regex.Replace(
                address,
                @"(?<=[a-zA-Z\u00C0-\u1EF90-9])(p\.?\d{1,2})(q\.?\d{1,2})\b",
                ", $1, $2",
                RegexOptions.IgnoreCase
            );
            address = Regex.Replace(
                address,
                @"(?<=\S)\s+(p\.?\s*\d{1,2})\b(?!\s*\d)",
                ", $1",
                RegexOptions.IgnoreCase
            );
            address = Regex.Replace(
                address,
                @"(?<=\S)\s+(q\.?\s*\d{1,2})\b",
                ", $1",
                RegexOptions.IgnoreCase
            );
            // Bước mới: chèn dấu phẩy trước "q.NAME" (quận không số), VD: "F22 . Q.bthanh" → "F22 . , Q.bthanh"
            // Pattern: trước q. phải có ký tự không-space, sau q. phải là chữ (không phải số)
            address = Regex.Replace(
                address,
                @"(?<=\S)[\s.]+((q|qu[aâậ]n)\.?\s*(?!\d)[a-zA-ZđÀ-ỹ][^\s,]*(?:\s+[^\s,]+)*)",
                ", $1",
                RegexOptions.IgnoreCase
            );
            address = Regex.Replace(
                address,
                @"(?<=\S)\s+(ph[uướừửữ][oôờ]ng\s+[^\s,]+(?:\s+[^\s,]+)*)",
                ", $1",
                RegexOptions.IgnoreCase
            );
            address = Regex.Replace(
                address,
                @"(?<=\S)\s+(qu[aâậ]n\s+[^\s,]+(?:\s+[^\s,]+)*)",
                ", $1",
                RegexOptions.IgnoreCase
            );

            foreach (var pattern in BuildKnownDistrictEndPatterns())
            {
                if (Regex.IsMatch(address, pattern, RegexOptions.IgnoreCase))
                {
                    address = Regex.Replace(address, pattern, ", $1", RegexOptions.IgnoreCase);
                    break;
                }
            }

            var segments = address
                .Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();

            int districtSegIdx = -1;
            for (int i = segments.Count - 1; i >= 0; i--)
            {
                var (quan, matchedDistText) = FindDistrictInSegment(segments[i]);
                if (!string.IsNullOrEmpty(quan))
                {
                    result.Quan = quan;
                    result.Confidence += 0.3f;
                    districtSegIdx = i;
                    // Nếu segment này bắt đầu bằng số (= có số nhà) VÀ có phần còn lại sau khi bỏ quận
                    // → chỉ strip phần quận ra, giữ lại phần còn lại làm địa chỉ
                    // VD: "208 Nguyễn Hữu Cảnh (22 Bình Thạnh ..." → strip "Bình Thạnh" → "208 Nguyễn Hữu Cảnh (22"
                    var seg = segments[i];
                    if (Regex.IsMatch(seg, @"^\d") && !string.IsNullOrEmpty(matchedDistText))
                    {
                        // Tìm vị trí của matchedDistText trong segment (bỏ phần đó và mọi thứ sau)
                        var stripped = Regex
                            .Replace(
                                seg,
                                @"[,\s]*" + Regex.Escape(matchedDistText) + @".*$",
                                "",
                                RegexOptions.IgnoreCase
                            )
                            .Trim();
                        if (!string.IsNullOrEmpty(stripped))
                        {
                            segments[i] = stripped; // Giữ lại, không xóa
                            districtSegIdx = -1; // Không xóa segment này
                        }
                    }
                    break;
                }
            }
            if (districtSegIdx >= 0)
                segments.RemoveAt(districtSegIdx);

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
            if (wardSegIdx >= 0)
                segments.RemoveAt(wardSegIdx);

            string remaining = string.Join(", ", segments).Trim();
            var (soNha, tenDuong) = ExtractHouseAndStreet(remaining);
            result.SoNha = soNha;
            result.TenDuong = CleanText(tenDuong);
            if (!string.IsNullOrEmpty(result.SoNha))
                result.Confidence += 0.2f;
            if (!string.IsNullOrEmpty(result.TenDuong))
                result.Confidence += 0.2f;
            result.Confidence = Math.Min(result.Confidence, 1f);

            result.Quan = NormalizeOutput(result.Quan);
            result.TenDuong = NormalizeOutput(result.TenDuong);
            result.SoNha = NormalizeOutput(result.SoNha);
            result.Phuong = NormalizeOutput(result.Phuong);

            // Nếu chưa tìm được quận nhưng có phường → thử tra bảng phường→quận
            if (string.IsNullOrEmpty(result.Quan) && !string.IsNullOrEmpty(result.Phuong))
            {
                var phuongNorm = NormalizeKey(result.Phuong);
                if (WardNormalizedDict.TryGetValue(phuongNorm, out var distFromWard))
                    result.Quan = distFromWard;
            }

            return result;
        }

        private static (string, string) FindDistrictInSegment(string seg)
        {
            seg = seg.Trim();
            if (string.IsNullOrEmpty(seg))
                return ("", "");

            var mQ = Regex.Match(seg, @"\bqu[aâậ]n\.?\s*(\d{1,2})\b", RegexOptions.IgnoreCase);
            if (!mQ.Success)
                mQ = Regex.Match(seg, @"\bq\.?\s*(\d{1,2})\b", RegexOptions.IgnoreCase);
            if (mQ.Success)
                return (mQ.Groups[1].Value, mQ.Value.Trim());

            // Xử lý "q.bthanh", "q tân bình", "Q.Bình Thạnh" — q./quận prefix + tên quận
            var mQName = Regex.Match(seg, @"^(?:qu[aâậ]n|q)\.?\s*(.+)$", RegexOptions.IgnoreCase);
            if (mQName.Success)
            {
                var nameAfterQ = mQName.Groups[1].Value.Trim();
                var nameNorm = NormalizeKey(nameAfterQ);
                if (DistrictNormalizedDict.TryGetValue(nameNorm, out var dQ))
                    return (dQ, seg);
                // Thử từng word trong phần sau q.
                var qWords = nameAfterQ.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = qWords.Length - 1; i >= 0; i--)
                {
                    if (DistrictNormalizedDict.TryGetValue(NormalizeKey(qWords[i]), out var d1))
                        return (d1, qWords[i]);
                    if (i > 0)
                    {
                        var pair = qWords[i - 1] + " " + qWords[i];
                        if (DistrictNormalizedDict.TryGetValue(NormalizeKey(pair), out var d2))
                            return (d2, pair);
                    }
                }
            }

            var segNorm = NormalizeKey(seg);
            if (DistrictNormalizedDict.TryGetValue(segNorm, out var dExact))
                return (dExact, seg);

            bool segIsJustNumber = Regex.IsMatch(seg, @"^\d{1,2}$");
            var words = seg.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = words.Length - 1; i >= 0; i--)
            {
                bool wordIsBarNumber = Regex.IsMatch(words[i], @"^\d{1,2}$");
                if (wordIsBarNumber && !segIsJustNumber)
                    continue;

                // Thử trực tiếp
                if (DistrictNormalizedDict.TryGetValue(NormalizeKey(words[i]), out var d1))
                    return (d1, words[i]);

                // Thử bỏ prefix q./quận khỏi từ, VD "Q.bthanh" → "bthanh"
                var stripped = Regex.Replace(
                    words[i],
                    @"^(qu[aâậ]n|q)\.?\s*",
                    "",
                    RegexOptions.IgnoreCase
                );
                if (stripped != words[i] && stripped.Length > 0)
                {
                    if (DistrictNormalizedDict.TryGetValue(NormalizeKey(stripped), out var dStrip))
                        return (dStrip, words[i]);
                }

                if (i > 0)
                {
                    var pair = words[i - 1] + " " + words[i];
                    if (DistrictNormalizedDict.TryGetValue(NormalizeKey(pair), out var d2))
                        return (d2, pair);
                }

                if (i > 1)
                {
                    var triple = words[i - 2] + " " + words[i - 1] + " " + words[i];
                    if (DistrictNormalizedDict.TryGetValue(NormalizeKey(triple), out var d3))
                        return (d3, triple);
                }
            }

            // Fallback: segment là tên phường → tra WardNormalizedDict để lấy quận trực tiếp
            // VD: segment = "an hội tây" → quận "go vap"
            var segNorm2 = NormalizeKey(seg);
            if (WardNormalizedDict.TryGetValue(segNorm2, out var dFromWard))
                return (dFromWard, seg);
            // Thử từng word pair/triple trong segment qua WardNormalizedDict
            for (int i = words.Length - 1; i >= 0; i--)
            {
                if (i > 0)
                {
                    var pair = words[i - 1] + " " + words[i];
                    if (WardNormalizedDict.TryGetValue(NormalizeKey(pair), out var dw2))
                        return (dw2, pair);
                }
                if (i > 1)
                {
                    var triple = words[i - 2] + " " + words[i - 1] + " " + words[i];
                    if (WardNormalizedDict.TryGetValue(NormalizeKey(triple), out var dw3))
                        return (dw3, triple);
                }
            }

            return ("", "");
        }

        private static (string, string) FindWardInSegment(string seg)
        {
            seg = seg.Trim();
            if (string.IsNullOrEmpty(seg))
                return ("", "");

            // Segment bắt đầu bằng số → đây là "số nhà + tên đường", không phải tên phường
            // VD: "200 lệ lại bến thành" — "bến thành" là địa danh trong đường, không phải phường
            if (Regex.IsMatch(seg, @"^\d"))
                return ("", "");

            foreach (var key in WardNameDict.Keys)
            {
                if (seg.Equals(key, StringComparison.OrdinalIgnoreCase))
                    return (WardNameDict[key], seg);
                if (seg.Contains(key, StringComparison.OrdinalIgnoreCase))
                    return (WardNameDict[key], key);
            }

            var mPNum = Regex.Match(
                seg,
                @"\b(?:ph[uướừửữ][oôờ]ng\.?\s*|phuong\.?\s*|p\.?\s*|f\.?\s*)(\d{1,2})\b",
                RegexOptions.IgnoreCase
            );
            if (mPNum.Success)
                return ($"Phường {mPNum.Groups[1].Value}", mPNum.Value.Trim());

            var segNoDiac = RemoveDiacriticsStatic(seg);
            if (segNoDiac.StartsWith("phuong") || segNoDiac.StartsWith("p."))
                return (seg, seg);

            // Thử tra WARD_TO_DISTRICT_MAP: nhận dạng tên phường chữ (VD: "Long Bình", "Linh Xuân")
            // Segment có thể là "phường long bình" hoặc chỉ "long bình"
            var segStripped = Regex
                .Replace(seg, @"^ph[uướừửữ][oôờ]ng\.?\s*", "", RegexOptions.IgnoreCase)
                .Trim();
            var segNorm = NormalizeKey(segStripped.Length > 0 ? segStripped : seg);
            if (WardNormalizedDict.ContainsKey(segNorm))
                return (segStripped.Length > 0 ? segStripped : seg, seg);

            // Thử từng word pair/triple trong segment
            var words = seg.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = words.Length - 1; i >= 0; i--)
            {
                if (i > 0)
                {
                    var pair = words[i - 1] + " " + words[i];
                    if (WardNormalizedDict.ContainsKey(NormalizeKey(pair)))
                        return (pair, pair);
                }
                if (i > 1)
                {
                    var triple = words[i - 2] + " " + words[i - 1] + " " + words[i];
                    if (WardNormalizedDict.ContainsKey(NormalizeKey(triple)))
                        return (triple, triple);
                }
            }

            return ("", "");
        }

        // Keyword nhận dạng đây là một tên đường (không phải landmark/mô tả)
        private static readonly Regex _streetKeywordRx = new Regex(
            @"^(?:(?:đường|duong|d\.|đ\.)\s+|(?:quốc\s*lộ|ql|tỉnh\s*lộ|tl|hẻm|hem|ngách|ngo|ngõ)\s*)[\w\d]",
            RegexOptions.IgnoreCase | RegexOptions.Compiled
        );

        private static (string soNha, string tenDuong) ExtractHouseAndStreet(string address)
        {
            var allSegs = address
                .Split(',')
                .Select(s => s.Trim())
                .Where(s => s.Length > 0)
                .ToList();
            var firstSeg = allSegs.Count > 0 ? allSegs[0] : address.Trim();

            // Khi có ≥2 segments và có segment bắt đầu bằng từ khóa đường
            // VD: "cổng số 2 ... chung cư khang gia, đường 45" → TÊN ĐƯỜNG = "đường 45", SỐ NHÀ = cả phần trước
            if (allSegs.Count >= 2)
            {
                int streetIdx = -1;
                for (int i = 0; i < allSegs.Count; i++)
                {
                    if (_streetKeywordRx.IsMatch(allSegs[i]))
                    {
                        streetIdx = i;
                        break;
                    }
                }
                if (streetIdx > 0) // street không phải segment đầu tiên
                {
                    var streetSeg = allSegs[streetIdx];
                    // Bỏ prefix "đường/d./đ." khỏi TÊN ĐƯỜNG nếu cần (giữ nguyên để normalize sau)
                    var streetName = Regex
                        .Replace(
                            streetSeg,
                            @"^(?:đường|duong|d\.|đ\.)\s*",
                            "",
                            RegexOptions.IgnoreCase
                        )
                        .Trim();
                    if (string.IsNullOrEmpty(streetName))
                        streetName = streetSeg;
                    // SỐ NHÀ = tất cả segments trước streetIdx (ghép lại bằng ", ")
                    var beforeStreet = string.Join(", ", allSegs.Take(streetIdx));
                    return (beforeStreet.Trim(), streetName);
                }
            }

            if (allSegs.Count >= 2 && Regex.IsMatch(firstSeg, @"^\d[\d\-/]*\d$"))
            {
                var nextSeg = allSegs[1];
                var streetMatch = Regex.Match(
                    nextSeg,
                    @"^(\d+(?:/\d+)?[A-Z]?)\s+(?:[\u0110\u0111]\.\s*)?(.+)$",
                    RegexOptions.IgnoreCase
                );
                if (streetMatch.Success)
                    return (firstSeg, streetMatch.Groups[2].Value.Trim());
                var streetName = Regex
                    .Replace(nextSeg, @"^\d+\s*[\u0110\u0111]\.\s*", "", RegexOptions.IgnoreCase)
                    .Trim();
                return (firstSeg, streetName.Length > 0 ? streetName : nextSeg);
            }

            var duongMatch = Regex.Match(
                firstSeg,
                @"^(?:s\u1ED1\s*)?(\d+(?:/\d+)?[A-Z]*)\s+(\u0111\u01B0\u1EDD\u006E\s+.+?)(?:\s+khu\b|\s+l\u00F4\b|\s+kdc\b|\s*$)",
                RegexOptions.IgnoreCase
            );
            if (duongMatch.Success)
                return (duongMatch.Groups[1].Value.Trim(), duongMatch.Groups[2].Value.Trim());

            var dotStreetMatch = Regex.Match(
                firstSeg,
                @"^(?:s\u1ED1\s*)?(\d+(?:/\d+)?[A-Z]*)\s+[\u0110\u0111]\.\s*(.+)$",
                RegexOptions.IgnoreCase
            );
            if (dotStreetMatch.Success)
                return (
                    dotStreetMatch.Groups[1].Value.Trim(),
                    dotStreetMatch.Groups[2].Value.Trim()
                );

            var numMatch = Regex.Match(
                firstSeg,
                @"^(.*\d+[A-Z]?)\s+([^\d(].*)$",
                RegexOptions.IgnoreCase
            );
            if (numMatch.Success)
                return (numMatch.Groups[1].Value.Trim(), numMatch.Groups[2].Value.Trim());

            var simpleMatch = Regex.Match(
                firstSeg,
                @"^(?:s\u1ED1\s*)?(\d+(?:/\d+)?[A-Z]*)\s*(.*)",
                RegexOptions.IgnoreCase
            );
            if (simpleMatch.Success)
                return (simpleMatch.Groups[1].Value.Trim(), simpleMatch.Groups[2].Value.Trim());

            // Không tìm được số nhà + tên đường từ pattern thông thường.
            // Kiểm tra nếu là dạng "landmark + số" (cổng 3, lô A, tầng 2, block D2...) → SỐ NHÀ = cả segment, TÊN ĐƯỜNG = ""
            // VD: "cong 3" → SỐ NHÀ="cong 3", "block d2" → SỐ NHÀ="block d2"
            if (
                Regex.IsMatch(
                    firstSeg,
                    @"^(?:cổng|cong|lô|lo|tầng|tang|block|căn|can|phòng|phong|kiosk|ki-ốt)\b",
                    RegexOptions.IgnoreCase
                )
            )
                return (firstSeg, "");

            return ("", firstSeg);
        }

        private static string NormalizeOutput(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return s;
            return RemoveDiacritics(s.Trim()).ToLowerInvariant();
        }

        private static string RemoveDiacritics(string s)
        {
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

        private static string RemoveDiacriticsStatic(string s)
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
            return sb.ToString().Normalize(System.Text.NormalizationForm.FormC).ToLowerInvariant();
        }

        private static string CleanText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";
            text = Regex.Replace(text, @"\s+", " ").Trim();
            text = text.TrimEnd(',', '.', ' ');
            return text;
        }

        public static string FormatAddress(ParsedAddress addr)
        {
            var parts = new List<string>();
            if (!string.IsNullOrEmpty(addr.SoNha))
                parts.Add(addr.SoNha);
            if (!string.IsNullOrEmpty(addr.TenDuong))
                parts.Add(addr.TenDuong);
            if (!string.IsNullOrEmpty(addr.Phuong))
                parts.Add(addr.Phuong);
            if (!string.IsNullOrEmpty(addr.Quan))
                parts.Add(addr.Quan);
            return string.Join(", ", parts.Where(p => !string.IsNullOrEmpty(p)));
        }
    }
}

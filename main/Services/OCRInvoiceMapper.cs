using System;
using System.Collections.Generic;

namespace TextInputter.Services
{
    /// <summary>
    /// Model: tất cả fields của 1 invoice (dùng bởi ExcelInvoiceService).
    /// </summary>
    public class OCRInvoiceData
    {
        public string Shop { get; set; } = "";
        public string TenKhachHang { get; set; } = "";
        public string SoHoaDon { get; set; } = "";
        public string DiaChi { get; set; } = "";
        public string Phuong { get; set; } = "";
        public string Quan { get; set; } = "";
        public decimal TongTienHang { get; set; } = 0;
        public decimal ChietKhau { get; set; } = 0;
        public decimal TongThanhToan { get; set; } = 0;
        public string NguoiDi { get; set; } = "";
        public string NguoiLay { get; set; } = "";
    }

    /// <summary>
    /// Helper tra cứu phí ship và người phụ trách theo địa chỉ.
    /// </summary>
    public class OCRInvoiceMapper
    {
        /// <summary>
        /// Tra cứu phí ship: ưu tiên theo PHƯỜNG (tier-3), fallback về QUẬN (tier-2).
        /// Trả về null nếu không tìm được → TIỀN SHIP để trống, user tự điền.
        /// </summary>
        /// <param name="phuong">Tên phường từ AddressParser (không dấu, lowercase)</param>
        /// <param name="quan">Tên quận từ AddressParser (không dấu, lowercase)</param>
        public static decimal? GetShipFee(string phuong, string quan)
        {
            // Tier-3: tra theo phường trong SHIPPING_FEES_BY_WARD (override giá quận)
            if (!string.IsNullOrWhiteSpace(phuong))
            {
                string normWard = NormalizeKey(phuong);
                foreach (var kv in AppConstants.SHIPPING_FEES_BY_WARD)
                {
                    if (NormalizeKey(kv.Key) == normWard)
                        return kv.Value;
                }

                // Tier-2.5: phường không có trong SHIPPING_FEES_BY_WARD
                // → thử tra WARD_TO_DISTRICT_MAP để ra quận, rồi tra ship theo quận đó
                foreach (var kv in AppConstants.WARD_TO_DISTRICT_MAP)
                {
                    if (NormalizeKey(kv.Key) == normWard)
                    {
                        // kv.Value là quận tương ứng — tra ship theo quận đó
                        string mappedQuan = NormalizeKey(kv.Value);
                        foreach (var sq in AppConstants.SHIPPING_FEES_BY_QUAN)
                        {
                            if (NormalizeKey(sq.Key) == mappedQuan)
                                return sq.Value;
                        }
                        break;
                    }
                }
            }

            // Tier-2: fallback về quận (normalize trước khi tra)
            if (!string.IsNullOrWhiteSpace(quan))
            {
                string normQuan = NormalizeKey(quan);
                foreach (var kv in AppConstants.SHIPPING_FEES_BY_QUAN)
                {
                    if (NormalizeKey(kv.Key) == normQuan)
                        return kv.Value;
                }
            }

            return null;
        }

        /// <summary>
        /// Backward-compat overload: chỉ tra theo quận (không có phường).
        /// </summary>
        public static decimal? GetShipFeeByQuan(string quan) => GetShipFee(null, quan);

        /// <summary>
        /// Tra cứu người phụ trách theo phường hoặc quận.
        /// Ưu tiên tra phường → fallback quận → fallback NGUOI_DI_DEFAULT.
        /// </summary>
        public static string GetNguoiDi(string phuong, string quan)
        {
            // Tra theo phường trước (AREA_TO_NGUOI_DI)
            if (!string.IsNullOrWhiteSpace(phuong))
            {
                string normWard = NormalizeKey(phuong);
                foreach (var kv in AppConstants.AREA_TO_NGUOI_DI)
                {
                    if (NormalizeKey(kv.Key) == normWard)
                        return kv.Value;
                }

                // Phường không có trong AREA_TO_NGUOI_DI → tra WARD_TO_DISTRICT_MAP → dùng quận
                foreach (var kv in AppConstants.WARD_TO_DISTRICT_MAP)
                {
                    if (NormalizeKey(kv.Key) == normWard)
                    {
                        string mappedQuan = NormalizeKey(kv.Value);
                        foreach (var aq in AppConstants.AREA_TO_NGUOI_DI)
                        {
                            if (NormalizeKey(aq.Key) == mappedQuan)
                                return aq.Value;
                        }
                        break;
                    }
                }
            }

            // Fallback quận (normalize trước khi tra)
            if (!string.IsNullOrWhiteSpace(quan))
            {
                string normQuan = NormalizeKey(quan);
                foreach (var kv in AppConstants.AREA_TO_NGUOI_DI)
                {
                    if (NormalizeKey(kv.Key) == normQuan)
                        return kv.Value;
                }
            }

            return AppConstants.NGUOI_DI_DEFAULT;
        }

        // Bảng expand viết tắt → tên đầy đủ (áp dụng trước khi tra cứu).
        // Thêm alias mới ở đây — không cần động vào AppConstants.
        private static readonly Dictionary<string, string> _abbrevMap = new Dictionary<
            string,
            string
        >(System.StringComparer.OrdinalIgnoreCase)
        {
            // Bình Thạnh
            { "bh thanh", "binh thanh" },
            { "b thanh", "binh thanh" },
            { "bthanh", "binh thanh" },
            { "b.thanh", "binh thanh" },
            // Tân Bình
            { "t binh", "tan binh" },
            { "tbinh", "tan binh" },
            { "t.binh", "tan binh" },
            // Tân Phú
            { "t phu", "tan phu" },
            { "tphu", "tan phu" },
            { "t.phu", "tan phu" },
            // Phú Nhuận
            { "p nhuan", "phu nhuan" },
            { "pnhuan", "phu nhuan" },
            { "p.nhuan", "phu nhuan" },
            // Gò Vấp
            { "g vap", "go vap" },
            { "gvap", "go vap" },
            { "g.vap", "go vap" },
            // Bình Tân
            { "b tan", "binh tan" },
            { "btan", "binh tan" },
            { "b.tan", "binh tan" },
            // Thủ Đức
            { "t duc", "thu duc" },
            { "tduc", "thu duc" },
            { "t.duc", "thu duc" },
            // Bình Chánh
            { "b chanh", "binh chanh" },
            { "bchanh", "binh chanh" },
        };

        private static string NormalizeKey(string s)
        {
            if (string.IsNullOrEmpty(s))
                return "";
            // 1. Strip diacritics + lowercase + giữ lại [a-z0-9 ]
            var noDiac = RemoveDiacritics(s);
            var norm = System
                .Text.RegularExpressions.Regex.Replace(noDiac.ToLowerInvariant(), @"[^a-z0-9 ]", "")
                .Trim();
            // 2. Expand viết tắt nếu match chính xác toàn chuỗi
            if (_abbrevMap.TryGetValue(norm, out string expanded))
                return expanded;
            return norm;
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
    }
}

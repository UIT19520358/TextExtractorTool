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
        public string SoNha { get; set; } = "";
        public string TenDuong { get; set; } = "";
        public string Phuong { get; set; } = "";
        public string Quan { get; set; } = "";
        public decimal TongTienHang { get; set; } = 0;
        public decimal ChietKhau { get; set; } = 0;
        public decimal TongThanhToan { get; set; } = 0;
        public string NguoiDi { get; set; } = "";
        public string NguoiLay { get; set; } = "";
    }

    /// <summary>
    /// Helper tra cứu phí ship theo quận.
    /// </summary>
    public class OCRInvoiceMapper
    {
        /// <summary>
        /// Tra cứu phí ship theo tên quận (output từ AddressParser — không dấu, lowercase).
        /// Trả về null nếu không tìm được → TIỀN SHIP để trống, user tự điền.
        /// </summary>
        public static decimal? GetShipFeeByQuan(string quan)
        {
            if (string.IsNullOrWhiteSpace(quan)) return null;

            if (AppConstants.SHIPPING_FEES_BY_QUAN.TryGetValue(quan.Trim(), out decimal fee))
                return fee;

            return null;
        }
    }
}

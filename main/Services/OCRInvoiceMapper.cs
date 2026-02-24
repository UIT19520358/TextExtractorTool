using System;
using System.Collections.Generic;

namespace TextInputter.Services
{
    /// <summary>
    /// Lớp để lưu dữ liệu hóa đơn từ OCR
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
    /// Lớp để map dữ liệu OCR sang Excel
    /// </summary>
    public class OCRInvoiceMapper
    {
        /// <summary>
        /// Map OCR invoice data sang Excel columns
        /// Trả về dictionary: Column Header -> Value
        /// </summary>
        public static Dictionary<string, string> MapToExcelColumns(OCRInvoiceData invoice)
        {
            var mapping = new Dictionary<string, string>
            {
                // Map giữa OCR fields và Excel column names
                // Điều chỉnh theo cấu trúc Excel thực tế của bạn
                { "MÃ", invoice.SoHoaDon },
                { "TÊN ĐƯ", invoice.TenDuong },
                { "SỐ NHÀ", invoice.SoNha },
                { "QUẬN", invoice.Quan },
                { "PHƯỜNG", invoice.Phuong },
                { "TIỀN HÀNG", invoice.TongTienHang.ToString("F0") },
                { "CHIẾT KHẤU", invoice.ChietKhau.ToString("F0") },
                { "THANH TOÁN", invoice.TongThanhToan.ToString("F0") },
                { "NGƯỜI ĐI", invoice.NguoiDi },
                { "NGƯỜI LẤY", invoice.NguoiLay },
            };

            return mapping;
        }

        /// <summary>
        /// Parse địa chỉ từ OCR, hiển thị dialog để user verify/edit
        /// </summary>
        public static (string soNha, string tenDuong, string phuong, string quan, bool success) 
            ParseAndVerifyAddress(string originalAddress)
        {
            // Parse tự động
            var parsed = AddressParser.Parse(originalAddress);

            // Nếu confidence cao (>= 0.7), tự động chấp nhận
            if (parsed.Confidence >= 0.7f)
            {
                return (parsed.SoNha, parsed.TenDuong, parsed.Phuong, parsed.Quan, true);
            }

            // Nếu confidence thấp, hiển thị dialog để user sửa
            var dialog = new AddressParsingDialog(
                originalAddress,
                parsed.SoNha,
                parsed.TenDuong,
                parsed.Phuong,
                parsed.Quan
            );

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return (dialog.SoNha, dialog.TenDuong, dialog.Phuong, dialog.Quan, true);
            }

            return ("", "", "", "", false);
        }
    }
}

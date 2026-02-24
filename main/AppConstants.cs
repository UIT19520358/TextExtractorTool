using System.Drawing;

namespace TextInputter
{
    /// <summary>
    /// Tất cả các giá trị hardcoded tập trung tại đây.
    /// Khi cần thay đổi config → chỉ sửa file này, không cần đụng logic.
    ///
    /// ⚠️  HARDCODED LIST (cần discuss để config-hoá sau):
    ///   - DAILY_REPORT_FILENAME   : tên file output cố định
    ///   - PHI_SHIP_MOI_DON        : 5đ/đơn — business rule, nên đọc từ config
    ///   - COL_SODON_FALLBACK_IDX  : index 17 (Column1) — phụ thuộc format Excel cụ thể
    ///   - HEADER_ROW_CANDIDATES   : "SHOP", "Tình trạng" — phụ thuộc file Excel của khách
    ///   - Google credential path  : tên file json cứng trong MainForm.cs
    /// </summary>
    internal static class AppConstants
    {
        // ── Paths ──────────────────────────────────────────────────────────────

        /// <summary>File output báo cáo hàng ngày, lưu cạnh .exe</summary>
        public const string DAILY_REPORT_FILENAME = "DailyTotalReport.xlsx";

        /// <summary>File Google Vision credentials</summary>
        public const string GOOGLE_CREDENTIAL_FILE = "textinputter-4a7bda4ef67a.json";

        /// <summary>Folder chứa ảnh OCR mặc định (dùng khi user chưa chọn)</summary>
        public const string DEFAULT_IMAGE_FOLDER = "";

        // ── Excel format ───────────────────────────────────────────────────────

        /// <summary>
        /// Các từ khoá để nhận ra hàng header trong Excel.
        /// ⚠️ HARDCODED: phụ thuộc format file Excel của khách hàng (Châu Ngân).
        /// Nếu đổi khách/template → cập nhật đây.
        /// </summary>
        public static readonly string[] HEADER_ROW_KEYWORDS = { "SHOP", "Tình trạng" };

        /// <summary>
        /// Scan tối đa N dòng đầu để tìm header row.
        /// </summary>
        public const int HEADER_SCAN_MAX_ROWS = 5;

        /// <summary>
        /// Index fallback của cột SỐ ĐƠN khi header không detect được.
        /// ⚠️ HARDCODED: cột "Column1" trong file Excel hiện tại nằm ở index 17.
        /// Nếu format Excel thay đổi → phải update.
        /// </summary>
        public const int COL_SODON_FALLBACK_IDX = 17;

        // ── Business rules ─────────────────────────────────────────────────────

        /// <summary>
        /// Phí ship thực tế mỗi đơn (VNĐ).
        /// ⚠️ HARDCODED: 5đ/đơn — nên đọc từ config file hoặc UI input sau.
        /// Công thức: khoanTruShip = -(totalTienShip - SoDon × PHI_SHIP_MOI_DON)
        /// </summary>
        public const decimal PHI_SHIP_MOI_DON = 5m;

        // ── UI colors ──────────────────────────────────────────────────────────

        /// <summary>Màu nền row TỔNG (vàng)</summary>
        public static readonly Color COLOR_ROW_TONG    = Color.Yellow;

        /// <summary>Màu nền row KẾT (vàng đậm)</summary>
        public static readonly Color COLOR_ROW_KET     = Color.FromArgb(255, 200, 0);

        /// <summary>Màu nền row âm / đơn trả (cam nhạt)</summary>
        public static readonly Color COLOR_ROW_NEGATIVE = Color.FromArgb(255, 200, 124);

        /// <summary>Màu nền dòng KẾT trong Daily Report panel (cam)</summary>
        public static readonly Color COLOR_REPORT_KET  = Color.FromArgb(255, 165, 0);

        /// <summary>Màu chữ đỏ cho số âm</summary>
        public static readonly Color COLOR_NEGATIVE_TEXT = Color.Red;

        /// <summary>Màu nền thanh button panel (dark)</summary>
        public static readonly Color COLOR_PANEL_DARK  = Color.FromArgb(40, 40, 40);

        // ── UI dimensions ──────────────────────────────────────────────────────

        /// <summary>Chiều cao panel Daily Report phía dưới</summary>
        public const int DAILY_REPORT_PANEL_HEIGHT = 220;

        /// <summary>Chiều cao row TỔNG trong dgvInvoice</summary>
        public const int ROW_HEIGHT_TONG = 24;

        /// <summary>Chiều cao row KẾT trong dgvInvoice</summary>
        public const int ROW_HEIGHT_KET  = 26;

        /// <summary>Chiều cao dòng KẾT trong Daily Report panel</summary>
        public const int ROW_HEIGHT_REPORT_KET = 28;

        // ── Date formats ───────────────────────────────────────────────────────

        /// <summary>Format ngày dùng làm sheet name khi Save</summary>
        public const string DATE_FORMAT_SHEET = "dd-MM-yyyy";

        /// <summary>Format ngày fallback (hôm nay)</summary>
        public const string DATE_FORMAT_DISPLAY = "dd.MM.yyyy";
    }
}

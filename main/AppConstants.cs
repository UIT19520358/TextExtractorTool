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

        /// <summary>
        /// Gemini API key — dùng để fallback parse địa chỉ khi AddressParser confidence thấp.
        /// Lấy miễn phí tại: https://aistudio.google.com/apikey
        /// Model fallback (quota nhiều → ít, tự động chuyển khi quota hết):
        ///   gemini-2.5-flash-lite → gemini-2.0-flash-lite → gemini-2.0-flash → gemini-2.5-flash → gemini-2.5-pro
        /// Để trống ("") = tắt Gemini fallback, chỉ dùng rule-based parser.
        /// ⚠️ Không commit key này lên git nếu repo public.
        /// </summary>
        public const string GEMINI_API_KEY = ""; // TODO: điền API key vào đây

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

        /// <summary>
        /// Người đi fallback — dùng khi quận/phường không có trong AREA_TO_NGUOI_DI.
        /// </summary>
        public const string NGUOI_DI_DEFAULT = "An Tam";

        // ── Excel / OCR defaults ────────────────────────────────────────────────

        /// <summary>
        /// Tên shop mặc định — fallback khi không tìm được SHOP từ OCR + Gemini.
        /// Thay đổi ở đây nếu tên shop thay đổi.
        /// </summary>
        public const string SHOP_DEFAULT = "ĐOÀN NGÂN CHÂU";

        /// <summary>
        /// Format ngày ghi vào Excel (cột NGÀY LẤY).
        /// VD: "28-02-2026." — có dấu chấm cuối theo quy ước của khách.
        /// Thay đổi ở đây nếu format ngày thay đổi.
        /// </summary>
        public const string DATE_FORMAT_EXCEL = "dd-MM-yyyy.";

        /// <summary>
        /// Màu nền đơn thất bại (FAIL) — dòng tô đỏ trong Excel.
        /// </summary>
        public const string COLOR_FAIL_ROW = "#FFD0D0";

        /// <summary>
        /// Màu nền đơn thiếu MÃ — tô đỏ đậm hơn để phân biệt cần tracking.
        /// </summary>
        public const string COLOR_MISSING_MA = "#FF9999";

        /// <summary>
        /// Regex pattern loại bỏ shopCandidate nếu trông như ghi chú thủ công.
        /// Thêm từ khoá vào đây khi có pattern ghi chú mới.
        /// </summary>
        public const string SHOP_EXCLUSION_PATTERN =
            @"nhận đổi|sản phẩm|ngày kể|giặt ủi|khui hàng|đổi size|"
            + @"không thể|không có|video|khiếu nại|giải quyết|mở hàng|"
            + @"lưu ý|ghi chú|chú ý|xin lỗi|cảm ơn";

        /// <summary>
        /// Người lấy cố định — người ở kho nhận/lấy hàng, không thay đổi theo khu vực.
        /// Thay đổi ở đây khi người lấy thay đổi.
        /// </summary>
        public const string NGUOI_LAY_DEFAULT = "c.cuong";

        /// <summary>
        /// Danh sách tất cả nhân viên — dùng cho dropdown ComboBox "Người Đi" / "Người Lấy".
        /// Thêm/xóa tên ở đây khi nhân sự thay đổi.
        /// </summary>
        public static readonly string[] NGUOI_LIST = new[]
        {
            "An Tam",
            "c.cuong",
            "c.hieu",
            "a.quyen",
        };

        /// <summary>
        /// Map phường/quận (key normalize không dấu) → tên người phụ trách khu vực đó.
        /// Key phải khớp giá trị QUẬN hoặc PHƯỜNG trả về từ AddressParser (không dấu, lowercase).
        /// Value: tên người đi (khớp với NGUOI_DI_DEFAULT hoặc tên khác sau này).
        ///
        /// Nếu QUẬN/PHƯỜNG không có trong map → dùng NGUOI_DI_DEFAULT làm fallback.
        /// ⚠️ HARDCODED: cập nhật khi phân công khu vực thay đổi.
        /// </summary>
        public static readonly Dictionary<string, string> AREA_TO_NGUOI_DI = new Dictionary<
            string,
            string
        >(System.StringComparer.OrdinalIgnoreCase)
        {
            // ── c.hiếu: Q1, Q3, Q4, Q5, Q10, Phú Nhuận ───────────────────────
            { "1", "c.hieu" },
            { "3", "c.hieu" },
            { "4", "c.hieu" },
            { "5", "c.hieu" },
            { "10", "c.hieu" },
            { "phu nhuan", "c.hieu" },
            { "tan binh", "c.hieu" },
            // ── c.cường: Bình Thạnh, Thủ Đức, Gò Vấp, Q2 ────────────────────
            { "binh thanh", "c.cuong" },
            // (alias "bh thanh", "b.thanh"... tự expand trong OCRInvoiceMapper._abbrevMap)
            { "thu duc", "c.cuong" },
            { "go vap", "c.cuong" },
            { "2", "c.cuong" },
            // ── a.quyền: Q9 ───────────────────────────────────────────────────
            { "9", "a.quyen" },

            // ── An Tâm: phần còn lại (Q6, Q7, Q8, Q11, Q12, Tân Phú, Bình Tân,
            //            Nhà Bè, Hóc Môn, Bình Chánh, Củ Chi, Cần Giờ...) ─────
            // → Handled by NGUOI_DI_DEFAULT = "An Tam"
            // Tân Bình: 1 phần → c.hiếu (ward-level needed), còn lại → An Tâm
            // TODO: khi AddressParser phân biệt được phường cụ thể của Tân Bình
            //       thì add ward-level mapping vào AREA_TO_NGUOI_DI.
            //       Hiện tại cả Tân Bình → An Tâm (fallback), c.hiếu nhận
            //       theo thỏa thuận nội bộ.
        };

        /// <summary>
        /// Bảng phí ship theo PHƯỜNG cụ thể — override bảng phí theo quận.
        /// Dùng khi cùng một quận nhưng các phường có phí ship khác nhau.
        ///
        /// Key  : tên phường không dấu, viết thường (NormalizeKey chuẩn hóa khi tra).
        ///         → Phải khớp key trong WARD_TO_DISTRICT_MAP hoặc giá trị Phuong từ AddressParser.
        /// Value: phí ship (nghìn đồng) — override SHIPPING_FEES_BY_QUAN.
        ///
        /// Ví dụ Q8 cũ có phường gần/xa phí khác nhau:
        ///   { "rach ong",   20m },  // gần → 20k
        ///   { "phu dinh",   30m },  // xa  → 30k
        ///
        /// ⚠️ HARDCODED: cập nhật khi bảng giá shipper thay đổi.
        /// Nếu phường không có trong bảng → fallback về SHIPPING_FEES_BY_QUAN theo quận.
        /// </summary>
        public static readonly Dictionary<string, decimal> SHIPPING_FEES_BY_WARD = new Dictionary<
            string,
            decimal
        >(System.StringComparer.OrdinalIgnoreCase)
        {
            // ── Quận 8 — nhóm 25k ─────────────────────────────────────────────
            // Phường mới (sau NQ1278/2024 + NQ1685/2025):
            { "rach ong", 25m }, // từ P.1+2+3 cũ
            { "hung phu", 25m }, // từ P.8+9+10 cũ
            { "chanh hung", 25m }, // từ P.4+phần P.5 cũ
            // Phường số cũ (trước 01/01/2025) — địa chỉ cũ vẫn dùng:
            { "phuong 1 quan 8", 25m },
            { "phuong 2 quan 8", 25m },
            { "phuong 3 quan 8", 25m },
            { "phuong 4 quan 8", 25m },
            { "phuong 8 quan 8", 25m },
            { "phuong 9 quan 8", 25m },
            { "phuong 10 quan 8", 25m },
            // ── Quận 8 — nhóm 30k ─────────────────────────────────────────────
            // Phường mới (sau NQ1685/2025):
            { "binh dong", 30m }, // từ P.6+7+phần P.5 cũ
            { "xom cui", 30m }, // từ P.11+12+13 cũ
            { "phu dinh", 30m }, // từ P.14+15+phần P.16 cũ
            { "hung thanh my", 30m },
            // Phường số cũ (trước 01/01/2025):
            { "phuong 5 quan 8", 30m },
            { "phuong 6 quan 8", 30m },
            { "phuong 7 quan 8", 30m },
            { "phuong 11 quan 8", 30m },
            { "phuong 12 quan 8", 30m },
            { "phuong 13 quan 8", 30m },
            { "phuong 14 quan 8", 30m },
            { "phuong 15 quan 8", 30m },
            { "phuong 16 quan 8", 30m },
        };

        /// <summary>
        /// Bảng phí ship theo quận/huyện (đơn vị: nghìn đồng, cùng đơn vị với TIỀN SHIP trong Excel).
        ///
        /// Key: đúng theo giá trị mà AddressParser trả về trong field QUẬN —
        ///   - Quận số → chỉ ghi số: "1", "2", ..., "12"
        ///   - Quận tên → dạng không dấu: "Binh Thanh", "Phu Nhuan", ...
        ///   GetShipFeeByQuan() sẽ tự normalize không dấu khi tra cứu,
        ///   nên key có dấu hay không dấu đều match được.
        ///
        /// Value: phí ship tính bằng nghìn đồng (k).
        ///   VD: 20m = 20 (nghìn đồng) = 20k, điền vào Excel cột TIỀN SHIP sẽ ra "20".
        ///   "m" là cú pháp C# cho kiểu decimal, không phải đơn vị khác.
        ///
        /// ⚠️ HARDCODED: phụ thuộc hợp đồng vận chuyển hiện tại.
        /// Khi đổi đơn vị vận chuyển / bảng giá → chỉ cần cập nhật đây.
        /// Nếu quận không có trong bảng → TIỀN SHIP để trống, user tự điền.
        /// </summary>
        public static readonly Dictionary<string, decimal> SHIPPING_FEES_BY_QUAN = new Dictionary<
            string,
            decimal
        >(System.StringComparer.OrdinalIgnoreCase)
        {
            // ── TP. HCM — quận số ─────────────────────────────────────────────
            // (key = số nguyên, khớp output AddressParser)
            { "1", 25m }, // Q1
            { "2", 30m }, // Q2
            { "3", 25m }, // Q3
            { "4", 25m }, // Q4
            { "5", 25m }, // Q5
            { "6", 25m }, // Q6
            { "7", 30m }, // Q7
            { "8", 25m }, // Q8 base 25k — các phường xa override lên 30k qua SHIPPING_FEES_BY_WARD
            { "9", 30m }, // Q9
            { "10", 25m }, // Q10
            { "11", 25m }, // Q11
            { "12", 30m }, // Q12
            // ── TP. HCM — quận/huyện tên ──────────────────────────────────────
            // (key lowercase không dấu, khớp output của AddressParser)
            { "binh thanh", 20m }, // Bình Thạnh
            // (alias "bh thanh", "b.thanh"... tự expand trong OCRInvoiceMapper._abbrevMap)
            { "phu nhuan", 20m }, // Phú Nhuận
            { "go vap", 25m }, // Gò Vấp
            { "tan binh", 25m }, // Tân Bình
            { "tan phu", 25m }, // Tân Phú
            { "binh tan", 30m }, // Bình Tân
            { "thu duc", 30m }, // Thủ Đức (TP.Thủ Đức cũ = Q2+Q9+Thủ Đức cũ, dùng 30k chung)
            // ── TP. HCM — huyện ngoại thành ───────────────────────────────────
            { "binh chanh", 35m }, // Bình Chánh
            { "hoc mon", 35m }, // Hóc Môn
            { "nha be", 35m }, // Nhà Bè
            { "cu chi", 40m }, // Củ Chi
            { "can gio", 50m }, // Cần Giờ
        };

        /// <summary>
        /// Map phường (tên duy nhất, không trùng quận) → quận cũ tương ứng.
        /// Dùng khi địa chỉ chỉ ghi phường mà không ghi quận (sau sáp nhập ĐVHC 2025).
        ///
        /// ⚠️ CHỈ map phường có tên DUY NHẤT (tên chữ, không phải "Phường 1/2/3...").
        ///   Phường số (Phường 1, 2, 3...) xuất hiện ở nhiều quận → KHÔNG map, bỏ qua.
        ///   Khi gặp phường số mà không có quận → TIỀN SHIP để trống, user tự điền.
        ///
        /// Key  : tên phường không dấu, viết thường (NormalizeKey chuẩn hoá khi tra).
        /// Value: quận cũ tương ứng — đúng format key của SHIPPING_FEES_BY_QUAN.
        ///
        /// Nguồn: Nghị quyết 1278/NQ-UBTVQH15 (sáp nhập ĐVHC TP.HCM 2025).
        /// Phí ship tính theo quận cũ, bảng giá giữ nguyên như cũ.
        ///
        /// TODO: bổ sung thêm khi có địa chỉ thực tế gặp phường chưa map
        /// </summary>
        public static readonly Dictionary<string, string> WARD_TO_DISTRICT_MAP = new Dictionary<
            string,
            string
        >(System.StringComparer.OrdinalIgnoreCase)
        {
            // ── Quận 1 (ship 20k) ─────────────────────────────────────────────
            // ⚠️ Quận 1 sau sáp nhập giữ nguyên tên "Quận 1", phường vẫn giữ tên cũ
            // Các phường số (Phường 1–10) trùng với quận khác → không map
            { "ben nghe", "1" }, // Phường Bến Nghé (tên duy nhất)
            { "ben thanh", "1" }, // Phường Bến Thành
            { "co giang", "1" }, // Phường Cô Giang
            { "cau kho", "1" }, // Phường Cầu Kho
            { "cau ong lanh", "1" }, // Phường Cầu Ông Lãnh
            { "da kao", "1" }, // Phường Đa Kao
            { "nguyen cu trinh", "1" }, // Phường Nguyễn Cư Trinh
            { "nguyen thai binh", "1" }, // Phường Nguyễn Thái Bình
            { "pham ngu lao", "1" }, // Phường Phạm Ngũ Lão
            { "tan dinh", "1" }, // Phường Tân Định
            // ── Quận 3 (ship 20k) ─────────────────────────────────────────────
            // Sau NQ1685/2025 (30/06/2025): Q3 còn 10 phường — phường số gộp thành tên mới
            { "vo thi sau", "3" }, // Phường Võ Thị Sáu — legacy (nay gộp vào Phường Xuân Hòa)
            { "ban co", "3" }, // Phường Bàn Cờ (mới 2025, từ P.1+2+3+5+phần P.4)
            { "xuan hoa", "3" }, // Phường Xuân Hòa (mới 2025, từ P.Võ Thị Sáu + phần P.4)
            { "nhieu loc", "3" }, // Phường Nhiêu Lộc (mới 2025, từ P.9+11+12+14)
            { "nguyen thai binh q3", "3" }, // trùng tên Q1 nên note thêm — không map tự động
            // ── Quận 4 (ship 25k) ─────────────────────────────────────────────
            // Sau NQ1685/2025 (01/07/2025): Q4 giải thể quận, còn 3 phường tên mới
            { "xom chieu", "4" }, // Phường Xóm Chiếu (mới 2025, từ P.13+16+18+phần P.15)
            { "khanh hoi", "4" }, // Phường Khánh Hội (mới 2025, từ P.8+9+phần P.2+4+15)
            { "vinh hoi", "4" }, // Phường Vĩnh Hội (mới 2025, từ P.1+3+phần P.2+4)
            { "tan thuan dong", "7" }, // Phường Tân Thuận Đông — Q7 (không phải Q4!)
            { "tan thuan tay", "7" }, // Phường Tân Thuận Tây — Q7
            // ── Quận 5 (ship 25k) ─────────────────────────────────────────────
            // Sau NQ1685/2025 (01/07/2025): Q5 giải thể quận, còn 3 phường tên mới
            { "cho lon", "5" }, // Phường Chợ Lớn (mới 2025, từ P.11+12+13+14) — ĐÃ CÓ ✅
            { "an dong", "5" }, // Phường An Đông (mới 2025, từ P.5+7+9)
            { "cho quan", "5" }, // Phường Chợ Quán (mới 2025, từ P.1+2+4)
            // ── Quận 10 (ship 25k) ────────────────────────────────────────────
            // Sau NQ1685/2025 (01/07/2025): Q10 giải thể quận, còn 3 phường tên mới
            { "vuon lai", "10" }, // Phường Vườn Lài (mới 2025, từ P.1+2+4+9+10)
            { "dien hong", "10" }, // Phường Diên Hồng (mới 2025, từ P.6+8+phần P.14)
            { "hoa hung", "10" }, // Phường Hòa Hưng (mới 2025, từ P.12+13+15+phần P.14)
            { "thanh thai", "10" }, // Đường Thành Thái đặc trưng Q10 (fallback địa chỉ cũ)
            // ── Quận Phú Nhuận (ship 20k) ─────────────────────────────────────
            { "phuong 17 phu nhuan", "phu nhuan" }, // dự phòng nếu OCR ghi rõ
            { "phu nhuan", "phu nhuan" },
            { "nguyen trong tuyen", "phu nhuan" }, // đường đặc trưng Q.PN — không chính xác
            // ── Quận Tân Bình (ship 25k) ──────────────────────────────────────
            // Sau NQ1685/2025 (30/06/2025): Tân Bình giải thể quận, còn 6 phường tên mới
            { "bay hien", "tan binh" }, // Phường Bảy Hiền (P.10+11+12 cũ)
            { "tan son hoa", "tan binh" }, // Phường Tân Sơn Hòa (mới 2025, từ P.1+2+3)
            { "tan son nhat", "tan binh" }, // Phường Tân Sơn Nhất (mới 2025, từ P.4+5+7)
            { "tan hoa q tan binh", "tan binh" }, // Phường Tân Hòa (mới 2025, từ P.6+8+9) — thêm hậu tố tránh trùng
            { "tan binh phuong", "tan binh" }, // Phường Tân Bình (mới 2025, từ P.13+14+phần P.15)
            { "tan son", "tan binh" }, // Phường Tân Sơn (mới 2025, từ phần còn lại P.15)
            // { "hoang van thu", "tan binh" }, // tên đường, không phải phường — loại bỏ
            // ── Quận Tân Phú (ship 25k) ───────────────────────────────────────
            { "tay thanh", "tan phu" }, // Phường Tây Thạnh
            { "son ky", "tan phu" }, // Phường Sơn Kỳ
            { "tan quy", "tan phu" }, // Phường Tân Quý
            { "tan son nhi", "tan phu" }, // Phường Tân Sơn Nhì
            { "tan thoi hoa", "tan phu" }, // Phường Tân Thới Hòa
            { "trang an", "tan phu" }, // Phường Trang An (duy nhất ở Tân Phú)
            { "hiep tan", "tan phu" }, // Phường Hiệp Tân
            { "hoa thanh", "tan phu" }, // Phường Hòa Thạnh
            // ── Quận Bình Thạnh (ship 20k) ────────────────────────────────────
            // Sau NQ1685/2025 (01/07/2025): Bình Thạnh giải thể quận, còn 5 phường tên mới
            { "binh thanh phuong", "binh thanh" }, // Phường Bình Thạnh (mới 2025, từ P.12+14+26) — hậu tố tránh trùng key quận
            { "gia dinh", "binh thanh" }, // Phường Gia Định (mới 2025, từ P.01+02+07+17)
            { "binh loi trung", "binh thanh" }, // Phường Bình Lợi Trung (mới 2025, từ P.05+11+13)
            { "thanh my tay", "binh thanh" }, // Phường Thạnh Mỹ Tây (mới 2025, từ P.19+22+25)
            { "binh quoi", "binh thanh" }, // Phường Bình Quới (mới 2025, từ P.27+28)
            { "hiep binh chanh", "thu duc" }, // Phường Hiệp Bình Chánh — TP.Thủ Đức cũ (không phải Bình Thạnh!)
            // TODO: Bình Thạnh sau sáp nhập phường vẫn số cũ → khó map

            // ── Quận Gò Vấp (ship 25k) ────────────────────────────────────────
            // Từ 01/07/2025 Gò Vấp giải thể quận → còn 6 phường tên (không còn phường số).
            // Tên cũ (trước 2025 — vẫn dùng khi OCR đọc địa chỉ cũ):
            { "hanh thong tay", "go vap" }, // Phường Hạnh Thông Tây cũ (nay thuộc P.Hạnh Thông)
            // Tên mới từ 01/07/2025:
            { "hanh thong", "go vap" }, // Phường Hạnh Thông (P1 cũ + P3 cũ)
            { "an nhon", "go vap" }, // Phường An Nhơn (P5 cũ + P6 cũ)
            { "go vap p", "go vap" }, // Phường Gò Vấp (P10 cũ + P17 cũ) — thêm "p" để tránh match "Quận Gò Vấp"
            { "thong tay hoi", "go vap" }, // Phường Thông Tây Hội (P8 cũ + P11 cũ)
            { "an hoi tay", "go vap" }, // Phường An Hội Tây (P12 cũ + P14 cũ)
            { "an hoi dong", "go vap" }, // Phường An Hội Đông (P15 cũ + P16 cũ)
            { "an hoi", "go vap" }, // Dự phòng nếu OCR đọc tắt "An Hội Tây/Đông"
            // ── Quận Bình Tân (ship 25k — TODO xác nhận) ──────────────────────
            // Sau NQ1685/2025: Bình Tân giải thể quận, còn 5 phường tên mới
            { "binh hung hoa", "binh tan" }, // Phường Bình Hưng Hòa (mới 2025)
            { "binh tri dong", "binh tan" }, // Phường Bình Trị Đông (mới 2025)
            { "tan tao", "binh tan" }, // Phường Tân Tạo (mới 2025, từ Tân Tạo A + phần Tân Tạo)
            { "an lac", "binh tan" }, // Phường An Lạc (mới 2025, từ An Lạc + An Lạc A + BT Đông B)
            // "tan tao a" → không còn tồn tại (gộp vào Phường Tân Tạo)
            // ── TP. Thủ Đức cũ (= Q2 + Q9 + Q.Thủ Đức cũ) ───────────────────
            // Ship chưa xác nhận, dùng "thu duc" làm key trung gian
            // Q2 cũ → ship 25k TODO | Q9 cũ → ship 30k TODO | Thủ Đức cũ → 25k TODO
            { "thu thiem", "2" }, // Phường Thủ Thiêm — Q2 cũ
            { "binh khanh", "2" }, // Phường Bình Khánh — Q2 cũ
            { "binh an", "2" }, // Phường Bình An — Q2 cũ
            { "an phu", "2" }, // Phường An Phú — Q2 cũ
            { "cat lai", "2" }, // Phường Cát Lái — Q2 cũ (nay Q9 mới?)
            { "long binh", "9" }, // Phường Long Bình — Q9 cũ → 30k
            { "long thanh my", "9" }, // Phường Long Thạnh Mỹ — Q9 cũ
            { "long phuoc", "9" }, // Phường Long Phước — Q9 cũ
            { "phuoc long a", "9" }, // Phường Phước Long A — Q9 cũ
            { "phuoc long b", "9" }, // Phường Phước Long B — Q9 cũ
            { "tang nhon phu a", "9" }, // Phường Tăng Nhơn Phú A — Q9 cũ
            { "tang nhon phu b", "9" }, // Phường Tăng Nhơn Phú B — Q9 cũ
            { "truong thanh", "9" }, // Phường Trường Thạnh — Q9 cũ
            { "phu huu", "9" }, // Phường Phú Hữu — Q9 cũ
            { "hiep phu", "9" }, // Phường Hiệp Phú — Q9 cũ
            { "linh xuan", "thu duc" }, // Phường Linh Xuân — Thủ Đức cũ
            { "linh dong", "thu duc" }, // Phường Linh Đông — Thủ Đức cũ
            { "linh chieu", "thu duc" }, // Phường Linh Chiểu — Thủ Đức cũ
            { "linh tay", "thu duc" }, // Phường Linh Tây — Thủ Đức cũ
            { "linh trung", "thu duc" }, // Phường Linh Trung — Thủ Đức cũ
            { "binh tho", "thu duc" }, // Phường Bình Thọ — Thủ Đức cũ
            { "binh chieu", "thu duc" }, // Phường Bình Chiểu — Thủ Đức cũ
            { "tam phu", "thu duc" }, // Phường Tam Phú — Thủ Đức cũ
            { "tam binh", "thu duc" }, // Phường Tam Bình — Thủ Đức cũ
            { "truong tho", "thu duc" }, // Phường Trường Thọ — Thủ Đức cũ
            { "hiep binh phuoc", "thu duc" }, // Phường Hiệp Bình Phước — Thủ Đức cũ
            // ── Quận 6 (ship 25k) ─────────────────────────────────────────────
            // Sau NQ1685/2025 (01/07/2025): Q6 giải thể quận, còn 4 phường tên mới
            { "binh tien", "6" }, // Phường Bình Tiên (mới 2025, từ P.1+7+8) — ĐÃ CÓ ✅
            { "binh tay", "6" }, // Phường Bình Tây (mới 2025, từ P.2+9)
            { "binh phu", "6" }, // Phường Bình Phú (mới 2025, từ P.10+11+phần P.16 Q8 cũ)
            { "phu lam", "6" }, // Phường Phú Lâm (mới 2025, từ P.12+13+14)
            // ── Quận 7 (ship 25k) ─────────────────────────────────────────────
            // Sau NQ1685/2025 (30/06/2025): Q7 giải thể quận, còn 4 phường tên mới
            { "tan my", "7" }, // Phường Tân Mỹ (mới 2025, từ P.Tân Phú+phần P.Phú Mỹ)
            { "tan hung", "7" }, // Phường Tân Hưng (mới 2025, từ P.Tân Phong+Tân Hưng+Tân Kiểng+Tân Quy)
            { "tan thuan", "7" }, // Phường Tân Thuận (mới 2025, từ P.Bình Thuận+Tân Thuận Đông+Tân Thuận Tây)
            { "phu thuan", "7" }, // Phường Phú Thuận (mới 2025, từ P.Phú Thuận+phần P.Phú Mỹ)
            { "phu my", "7" }, // Phường Phú Mỹ cũ — Q7 (legacy, nay thuộc Tân Mỹ/Phú Thuận)
            { "phuoc kien", "7" }, // Phước Kiển — Q7 (legacy)
            { "hung gia", "7" }, // khu Hưng Gia, Phú Mỹ Hưng — Q7
            { "hung phuoc", "7" }, // khu Hưng Phước — Q7
            { "tan phu q7", "7" }, // Phường Tân Phú cũ — Q7 (trùng Quận Tân Phú → bỏ qua)
            // ── Quận 8 (ship 25k) ─────────────────────────────────────────────
            // Từ NQ1278/2024 (01/01/2025): phường số gộp thành Rạch Ông, Hưng Phú, Xóm Củi
            // Từ NQ1685/2025 (01/07/2025): Q8 giải thể quận, còn 10 phường tên mới
            { "binh dong", "8" }, // Phường Bình Đông (mới 2025)
            { "chanh hung", "8" }, // Phường Chánh Hưng (mới 2025, từ Rạch Ông+Hưng Phú+P.4+phần P.5)
            { "phu dinh", "8" }, // Phường Phú Định (mới 2025, từ Xóm Củi+P.14+15+phần P.16)
            { "rach ong", "8" }, // Phường Rạch Ông (từ NQ1278/2024: P.1+2+3)
            { "hung phu", "8" }, // Phường Hưng Phú (từ NQ1278/2024: P.8+9+10)
            { "xom cui", "8" }, // Phường Xóm Củi (từ NQ1278/2024: P.11+12+13)
            { "hung thanh my", "8" }, // Phường Hưng Thạnh Mỹ — Q8
            // ── Quận 11 (ship 25k) ────────────────────────────────────────────
            // Sau NQ1685/2025 (30/06/2025): Q11 giải thể quận, còn 4 phường tên mới
            { "minh phung", "11" }, // Phường Minh Phụng (mới 2025, từ P.1+7+16)
            { "binh thoi", "11" }, // Phường Bình Thới (mới 2025, từ P.3+phần P.8)
            { "phu tho q11", "11" }, // Phường Phú Thọ (mới 2025, từ P.11+15+phần P.8) — hậu tố tránh trùng
            { "hoa binh q11", "11" }, // Phường Hòa Bình (mới 2025, từ P.5+14) — hậu tố tránh trùng
            // ── Quận 12 (ship 25k) ────────────────────────────────────────────
            { "thanh loc", "12" }, // Phường Thạnh Lộc
            { "thanh xuan", "12" }, // Phường Thạnh Xuân
            { "trung my tay", "12" }, // Phường Trung Mỹ Tây
            { "tan thoi nhat", "12" }, // Phường Tân Thới Nhất
            { "tan chanh hiep", "12" }, // Phường Tân Chánh Hiệp
            { "dong hung thuan", "12" }, // Phường Đông Hưng Thuận
            { "thoi an", "12" }, // Phường Thới An
            { "hiep thanh", "12" }, // Phường Hiệp Thành — Q12
        };

        // ── UI colors ──────────────────────────────────────────────────────────

        /// <summary>Màu nền row TỔNG (vàng)</summary>
        public static readonly Color COLOR_ROW_TONG = Color.Yellow;

        /// <summary>Màu nền row KẾT (vàng đậm)</summary>
        public static readonly Color COLOR_ROW_KET = Color.FromArgb(255, 200, 0);

        /// <summary>Màu nền row âm / đơn trả (cam nhạt)</summary>
        public static readonly Color COLOR_ROW_NEGATIVE = Color.FromArgb(255, 200, 124);

        /// <summary>Màu nền dòng KẾT trong Daily Report panel (cam)</summary>
        public static readonly Color COLOR_REPORT_KET = Color.FromArgb(255, 165, 0);

        /// <summary>Màu chữ đỏ cho số âm</summary>
        public static readonly Color COLOR_NEGATIVE_TEXT = Color.Red;

        /// <summary>Màu nền thanh button panel (dark)</summary>
        public static readonly Color COLOR_PANEL_DARK = Color.FromArgb(40, 40, 40);

        // ── UI dimensions ──────────────────────────────────────────────────────

        /// <summary>Chiều cao panel Daily Report phía dưới</summary>
        public const int DAILY_REPORT_PANEL_HEIGHT = 220;

        /// <summary>Chiều cao row TỔNG trong dgvInvoice</summary>
        public const int ROW_HEIGHT_TONG = 24;

        /// <summary>Chiều cao row KẾT trong dgvInvoice</summary>
        public const int ROW_HEIGHT_KET = 26;

        /// <summary>Chiều cao dòng KẾT trong Daily Report panel</summary>
        public const int ROW_HEIGHT_REPORT_KET = 28;

        // ── Date formats ───────────────────────────────────────────────────────

        /// <summary>Format ngày dùng làm sheet name khi Save</summary>
        public const string DATE_FORMAT_SHEET = "dd-MM-yyyy";

        /// <summary>Format ngày fallback (hôm nay)</summary>
        public const string DATE_FORMAT_DISPLAY = "dd.MM.yyyy";
    }
}

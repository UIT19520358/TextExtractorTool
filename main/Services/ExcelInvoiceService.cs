using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace TextInputter.Services
{
    /// <summary>
    /// Service ghi dữ liệu invoice vào file Excel của khách.
    /// File path được truyền vào từ caller (qua OpenFileDialog) — không hardcode tên file.
    /// </summary>
    public class ExcelInvoiceService
    {
        private readonly string _excelFilePath;

        /// <param name="excelFilePath">Full path đến file Excel (lấy từ OpenFileDialog ở caller).</param>
        /// <exception cref="FileNotFoundException">Nếu file không tồn tại.</exception>
        public ExcelInvoiceService(string excelFilePath)
        {
            if (!File.Exists(excelFilePath))
                throw new FileNotFoundException($"Excel file not found: {excelFilePath}");
            _excelFilePath = excelFilePath;
        }

        /// <summary>
        /// Kiểm tra hóa đơn với cùng số có tồn tại chưa.
        /// </summary>
        public bool InvoiceExists(string soHoaDon, out string existingSheet)
        {
            existingSheet = null;

            try
            {
                using (var workbook = new XLWorkbook(_excelFilePath))
                {
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        var rows = worksheet.RowsUsed();
                        foreach (var row in rows)
                        {
                            if (row.RowNumber() <= 2)
                                continue; // Skip header rows (row1=cols, row2=THU x)

                            var cell = row.Cell(COL_MA); // MÃ column (invoice number)
                            string cellValue = cell.GetString();
                            if (cellValue == soHoaDon)
                            {
                                existingSheet = worksheet.Name;
                                return true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error checking invoice: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// Lấy tất cả số hóa đơn đã có trong Excel.
        /// </summary>
        public List<string> GetAllInvoiceNumbers()
        {
            var invoices = new List<string>();

            try
            {
                using (var workbook = new XLWorkbook(_excelFilePath))
                {
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        var rows = worksheet.RowsUsed();
                        foreach (var row in rows)
                        {
                            if (row.RowNumber() <= 2)
                                continue; // Skip header rows (row1=cols, row2=THU x)

                            var cell = row.Cell(COL_MA); // MÃ column (invoice number)
                            string cellValue = cell.GetString();
                            if (!string.IsNullOrWhiteSpace(cellValue))
                                invoices.Add(cellValue);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting invoices: {ex.Message}");
            }

            return invoices;
        }

        // ── Column index constants (1-based) ────────────────────────────────────
        // Thay đổi ở đây nếu cấu trúc cột Excel thay đổi.
        private const int COL_TINHTRANG = 1;
        private const int COL_SHOP = 2;
        private const int COL_TENKH = 3;
        private const int COL_MA = 4;
        private const int COL_DIACHI = 5; // gộp SỐ NHÀ + TÊN ĐƯỜNG
        private const int COL_QUAN = 6;
        private const int COL_TIENTHU = 7;
        private const int COL_TIENSHIP = 8;
        private const int COL_TIENHANG = 9;
        private const int COL_NGUOIDI = 10;
        private const int COL_NGUOILAY = 11;
        private const int COL_NGAYLAY = 12;
        private const int COL_GHICHU = 13;
        private const int COL_UNGIEN = 14;
        private const int COL_HANGTON = 15;
        private const int COL_FAIL = 16;
        private const int COL_COL1 = 17;
        private const int COL_COL2 = 18;
        private const int COL_COL3 = 19;

        // ── Layout constants ─────────────────────────────────────────────────────
        // Data bắt đầu từ row 3 (row 1+2 là header).
        private const int DATA_START_ROW = 3;

        // Số dòng mỗi block bảng tổng kết bên phải (per NGƯỜI ĐI).
        private const int SUMMARY_RIGHT_BLOCK_HEIGHT = 7;

        // Số dòng mỗi block bảng tổng kết bên trái (per SHOP).
        private const int SUMMARY_LEFT_BLOCK_HEIGHT = 8;

        // Số dòng dự phòng khi xóa summary cũ (tránh bỏ sót).
        private const int SUMMARY_CLEAR_EXTRA_ROWS = 50;

        /// <summary>
        /// Xuất nhiều dòng cùng lúc (batch) vào một sheet — chỉ ghi data rows thuần,
        /// KHÔNG add formula, SUBTOTAL hay bảng tổng kết.
        /// Formula/summary được add riêng bởi ApplyFormulasAndSummary() sau khi user bấm "Tính Tiền".
        /// </summary>
        /// <param name="dataList">Danh sách dict với keys: SHOP, TÊN KH, MÃ, ĐỊA CHỈ, QUẬN,
        ///     TIỀN THU, TIỀN SHIP, NGƯỜI ĐI, NGƯỜI LẤY, NGÀY LẤY.</param>
        /// <param name="sheetName">Tên sheet (vd: "25-07"). Nếu chưa có → tự tạo với header row.</param>
        /// <param name="sheetDate">Ngày dùng để tạo header row 2 (THU x / NGAY x-x).</param>
        /// <returns>(addedCount, updatedCount)</returns>
        public (int added, int updated) ExportBatch(
            IEnumerable<Dictionary<string, string>> dataList,
            string sheetName,
            DateTime sheetDate
        )
        {
            int addedCount = 0,
                updatedCount = 0;

            using (var workbook = new XLWorkbook(_excelFilePath))
            {
                IXLWorksheet worksheet;
                if (workbook.TryGetWorksheet(sheetName, out var existingSheet))
                {
                    worksheet = existingSheet;
                }
                else
                {
                    worksheet = workbook.Worksheets.Add(sheetName);
                    AddHeaderRow(worksheet, sheetDate);
                }

                // Tìm SUBTOTAL row (nếu đã có) — KHÔNG ghi đè vào đó, chỉ ghi đè khi MÃ trùng
                int existingSubtotalRow = FindSubtotalRow(worksheet);

                // Next empty data row = hàng trước SUBTOTAL (nếu có) hoặc sau last used row.
                // Nếu SUBTOTAL tồn tại, ta cần INSERT rows để đẩy SUBTOTAL xuống thay vì overwrite.
                int nextRow = DATA_START_ROW;
                if (existingSubtotalRow > 0)
                {
                    // Đếm số đơn mới thực sự (không phải upsert) để biết cần chèn bao nhiêu row
                    int newRowsNeeded = dataList.Count(d =>
                    {
                        string mCheck = d.GetValueOrDefault("MÃ", "");
                        if (string.IsNullOrWhiteSpace(mCheck))
                        {
                            // MÃ rỗng → kiểm tra có row khớp TÊN KH + NGÀY LẤY không
                            string tenKhCheck = d.GetValueOrDefault("TÊN KH", "");
                            string ngayLayCheck = d.GetValueOrDefault("NGÀY LẤY", "");
                            return FindRowByTenKhNgay(worksheet, tenKhCheck, ngayLayCheck) < 0; // không tìm thấy → thêm mới
                        }
                        return FindRowByMa(worksheet, mCheck) < 0; // không tìm thấy → thêm mới
                    });
                    // Chèn đủ số hàng trống trước SUBTOTAL row
                    if (newRowsNeeded > 0)
                        worksheet.Row(existingSubtotalRow).InsertRowsAbove(newRowsNeeded);
                    // nextRow bắt đầu từ vị trí trước khi chèn (SUBTOTAL đã dịch xuống)
                    nextRow = existingSubtotalRow;
                }
                else
                {
                    var lastUsed = worksheet.LastRowUsed();
                    if (lastUsed != null && lastUsed.RowNumber() >= DATA_START_ROW)
                        nextRow = lastUsed.RowNumber() + 1;
                }

                foreach (var data in dataList)
                {
                    string ma = data.GetValueOrDefault("MÃ", "");
                    bool isMissing = string.IsNullOrWhiteSpace(ma);
                    bool hasMissingFields = !string.IsNullOrEmpty(
                        data.GetValueOrDefault("MISSING_FIELDS", "")
                    );

                    // Upsert: tìm row có MÃ trùng → ghi đè; nếu MÃ rỗng → tìm theo TÊN KH + NGÀY LẤY
                    // (đơn hàng sỉ không có MÃ HĐ); nếu vẫn không thấy → append mới
                    int targetRow;
                    if (!isMissing)
                    {
                        targetRow = FindRowByMa(worksheet, ma);
                    }
                    else
                    {
                        string tenKh = data.GetValueOrDefault("TÊN KH", "");
                        string ngayLay = data.GetValueOrDefault("NGÀY LẤY", "");
                        targetRow = FindRowByTenKhNgay(worksheet, tenKh, ngayLay);
                    }
                    bool isUpdate = targetRow > 0;
                    if (!isUpdate)
                    {
                        targetRow = nextRow;
                        nextRow++;
                    }

                    // Nếu đơn có field thiếu và SHOP trông như câu ghi chú,
                    // chuyển giá trị SHOP sang GHI CHÚ để không ô nhiễm cột SHOP.
                    string shopVal = data.GetValueOrDefault("SHOP", "");
                    string ghichuVal = data.GetValueOrDefault("GHI CHÚ", "");
                    if (hasMissingFields && ShopLooksLikeNote(shopVal))
                    {
                        ghichuVal = string.IsNullOrWhiteSpace(ghichuVal)
                            ? shopVal.Trim()
                            : shopVal.Trim() + " | " + ghichuVal;
                        shopVal = "";
                    }

                    // Ghi data vào hàng (isFail=false vì không còn phân biệt fail/success)
                    WriteDataRow(
                        worksheet,
                        targetRow,
                        data,
                        shopVal,
                        ghichuVal,
                        ma,
                        false,
                        isMissing
                    );

                    if (isUpdate)
                        updatedCount++;
                    else
                        addedCount++;
                }

                workbook.SaveAs(_excelFilePath);
            }

            return (addedCount, updatedCount);
        }

        /// <summary>
        /// Thêm TIỀN HÀNG formula, SUBTOTAL row và bảng tổng kết (per shop + per người đi)
        /// vào sheet đã có data. Gọi sau khi user bấm "Tính Tiền" ở tab Excel Viewer.
        /// Nếu sheet đã có SUBTOTAL/summary → xóa và viết lại.
        /// </summary>
        /// <param name="sheetName">Tên sheet cần xử lý.</param>
        /// <param name="sheetDate">Ngày của sheet (dùng khi không tìm được ngày trong data).</param>
        public void ApplyFormulasAndSummary(string sheetName, DateTime sheetDate)
        {
            using (var workbook = new XLWorkbook(_excelFilePath))
            {
                if (!workbook.TryGetWorksheet(sheetName, out var worksheet))
                    throw new InvalidOperationException($"Sheet '{sheetName}' không tồn tại.");

                // ── Xóa SUBTOTAL + summary cũ nếu có ──────────────────────────
                int existingSubtotalRow = -1;
                int existingSummaryStartRow = -1;
                foreach (var row in worksheet.RowsUsed().ToList())
                {
                    int rn = row.RowNumber();
                    if (rn < DATA_START_ROW)
                        continue;
                    var cell = row.Cell(COL_TIENTHU);
                    if (
                        cell.HasFormula
                        && cell.FormulaA1.StartsWith("SUBTOTAL", StringComparison.OrdinalIgnoreCase)
                    )
                    {
                        existingSubtotalRow = rn;
                    }
                    if (
                        existingSubtotalRow > 0
                        && existingSummaryStartRow < 0
                        && cell.HasFormula
                        && cell.FormulaA1.StartsWith("SUMIFS", StringComparison.OrdinalIgnoreCase)
                    )
                    {
                        existingSummaryStartRow = rn;
                    }
                }
                if (existingSubtotalRow > 0)
                {
                    int clearFrom = existingSubtotalRow;
                    int clearTo =
                        existingSummaryStartRow > 0
                            ? existingSummaryStartRow + SUMMARY_CLEAR_EXTRA_ROWS
                            : existingSubtotalRow + 1;
                    for (int r = clearFrom; r <= clearTo; r++)
                        worksheet.Row(r).Clear();
                }

                // ── Tìm last data row ──────────────────────────────────────────
                // Chỉ tính row có MÃ (cột D) — tránh bị summary cũ làm sai lastDataRow
                int lastDataRow = DATA_START_ROW - 1;
                foreach (var row in worksheet.RowsUsed())
                {
                    int rn = row.RowNumber();
                    if (rn < DATA_START_ROW)
                        continue;
                    string maVal = row.Cell(COL_MA).GetString().Trim();
                    string shopVal = row.Cell(COL_SHOP).GetString().Trim();
                    // Bỏ qua row header
                    if (shopVal.Equals("SHOP", StringComparison.OrdinalIgnoreCase))
                        continue;
                    // Chỉ tính row có MÃ ĐƠN (data thật), hoặc có SHOP nhưng ko có MÃ (đơn hàng si/hoàn)
                    // Bỏ qua row từ summary cũ (có text như "cod", "Trừ Tiền Ship", "THANH TOÁN" trong COL_SHOP nhưng ko có MÃ)
                    if (string.IsNullOrWhiteSpace(maVal))
                    {
                        // Không có MÃ → chỉ tính nếu SHOP là tên shop thật (không phải text summary)
                        if (string.IsNullOrWhiteSpace(shopVal))
                            continue;
                        string shopLower = shopVal.ToLower();
                        if (
                            shopLower == "cod"
                            || shopLower.StartsWith("trừ")
                            || shopLower.StartsWith("tru")
                            || shopLower.StartsWith("tiền")
                            || shopLower.StartsWith("tien")
                            || shopLower == "đơn đơn"
                            || shopLower.StartsWith("don")
                            || shopLower == "nợ cũ"
                            || shopLower.StartsWith("no ")
                            || shopLower == "thanh toán"
                            || shopLower == "thanh toan"
                            || shopLower == "hàng si"
                            || shopLower == "hang si"
                        )
                            continue; // đây là summary row cũ
                    }
                    lastDataRow = rn;
                }
                if (lastDataRow < DATA_START_ROW)
                    throw new InvalidOperationException(
                        $"Sheet '{sheetName}' không có dữ liệu đơn hàng (từ row {DATA_START_ROW} trở đi).\n"
                            + "Vui lòng chọn file Excel đã có data nhập vào (cùng format với file gốc)."
                    );

                // ── Fix TIỀN HÀNG formula cho data rows (ghi đè = đảm bảo đúng) ──
                string thuColLetter = ColLetter(COL_TIENTHU);
                string shipColLetter = ColLetter(COL_TIENSHIP);
                for (int r = DATA_START_ROW; r <= lastDataRow; r++)
                {
                    // Bỏ qua hàng hoàn toàn trống (không có MÃ, SHOP, hay TIỀN THU)
                    bool rowHasData =
                        !string.IsNullOrWhiteSpace(worksheet.Cell(r, COL_MA).GetString())
                        || !string.IsNullOrWhiteSpace(worksheet.Cell(r, COL_SHOP).GetString())
                        || !string.IsNullOrWhiteSpace(worksheet.Cell(r, COL_TIENTHU).GetString());
                    if (!rowHasData)
                        continue;

                    var hangCell = worksheet.Cell(r, COL_TIENHANG);
                    // Ghi đè formula COD (=G+H). Giữ nguyên nếu đã là SHIP_ONLY (=-H hoặc =+H thuần)
                    string existingFormula = hangCell.HasFormula ? hangCell.FormulaA1.Trim() : "";
                    bool isShipOnlyFormula =
                        existingFormula == $"-{shipColLetter}{r}"
                        || existingFormula == $"{shipColLetter}{r}";
                    if (!isShipOnlyFormula)
                        hangCell.FormulaA1 = $"{thuColLetter}{r}-{shipColLetter}{r}"; // COD: thu - ship
                    hangCell.Style.NumberFormat.Format = "#,##0";
                }

                // ── SUBTOTAL row ───────────────────────────────────────────────
                int subtotalRow = lastDataRow + 2;
                worksheet.Row(lastDataRow + 1).Clear();

                // Chỉ SUBTOTAL các cột số thực sự — bỏ cột text (NGƯỜI ĐI, NGÀY LẤY, GHI CHÚ, v.v.)
                int[] subtotalCols =
                {
                    COL_TIENTHU,
                    COL_TIENSHIP,
                    COL_TIENHANG,
                    COL_UNGIEN,
                    COL_HANGTON,
                };
                foreach (int col in subtotalCols)
                {
                    string colLetter = ColLetter(col);
                    var stCell = worksheet.Cell(subtotalRow, col);
                    stCell.FormulaA1 =
                        $"SUBTOTAL(9,{colLetter}{DATA_START_ROW}:{colLetter}{lastDataRow})";
                    stCell.Style.Font.Bold = true;
                    stCell.Style.Fill.BackgroundColor = XLColor.LightYellow;
                    stCell.Style.NumberFormat.Format = "#,##0";
                }

                // ── Bảng tổng kết ─────────────────────────────────────────────
                int summaryRow = subtotalRow + 2;

                // Clear row giữa SUBTOTAL và summary để tránh dư data cũ
                worksheet.Row(subtotalRow + 1).Clear();

                // Thu thập distinct SHOPs, NGƯỜI ĐIs và ngày đầu tiên từ data rows
                var distinctShops = new List<string>();
                var distinctNguoiDis = new List<string>();
                string firstNgay = "";
                for (int r = DATA_START_ROW; r <= lastDataRow; r++)
                {
                    string shop = worksheet.Cell(r, COL_SHOP).GetString().Trim();
                    string nguoiDi = worksheet.Cell(r, COL_NGUOIDI).GetString().Trim();
                    string ngay = worksheet.Cell(r, COL_NGAYLAY).GetString().Trim();
                    if (!string.IsNullOrWhiteSpace(shop) && !distinctShops.Contains(shop))
                        distinctShops.Add(shop);
                    if (!string.IsNullOrWhiteSpace(nguoiDi) && !distinctNguoiDis.Contains(nguoiDi))
                        distinctNguoiDis.Add(nguoiDi);
                    if (string.IsNullOrWhiteSpace(firstNgay) && !string.IsNullOrWhiteSpace(ngay))
                        firstNgay = ngay;
                }
                if (distinctShops.Count == 0)
                    distinctShops.Add(AppConstants.SHOP_DEFAULT);
                if (string.IsNullOrWhiteSpace(firstNgay))
                    firstNgay = sheetDate.ToString(AppConstants.DATE_FORMAT_EXCEL);

                // Shorthand cột letters (dùng nhiều trong SUMIFS/COUNTIFS)
                string shopColL = ColLetter(COL_SHOP);
                string ngayColL = ColLetter(COL_NGAYLAY);
                string thuColL = ColLetter(COL_TIENTHU);
                string shipColL = ColLetter(COL_TIENSHIP);
                string nguoiDiColL = ColLetter(COL_NGUOIDI);
                string tenkhColL = ColLetter(COL_TENKH);
                string diachiColL = ColLetter(COL_DIACHI);
                string maColL = ColLetter(COL_MA);

                // SUMIFS/COUNTIFS ranges cố định vào vùng data
                string rThu = $"{thuColL}${DATA_START_ROW}:{thuColL}${lastDataRow}";
                string rShip = $"{shipColL}${DATA_START_ROW}:{shipColL}${lastDataRow}";
                string rShop = $"{shopColL}${DATA_START_ROW}:{shopColL}${lastDataRow}";
                string rNgay = $"{ngayColL}${DATA_START_ROW}:{ngayColL}${lastDataRow}";
                string rNguoiDi = $"{nguoiDiColL}${DATA_START_ROW}:{nguoiDiColL}${lastDataRow}";
                string rMa = $"{maColL}${DATA_START_ROW}:{maColL}${lastDataRow}";

                // ── BẢNG PHẢI: per NGƯỜI ĐI ───────────────────────────────────
                BuildRightSummary(
                    worksheet,
                    distinctNguoiDis,
                    summaryRow,
                    rThu,
                    rShip,
                    rNguoiDi,
                    rShop
                );

                // ── BẢNG TRÁI: per SHOP ───────────────────────────────────────
                BuildLeftSummary(
                    worksheet,
                    distinctShops,
                    summaryRow,
                    firstNgay,
                    rThu,
                    rShip,
                    rShop,
                    rNgay,
                    rMa,
                    tenkhColL,
                    diachiColL
                );

                workbook.SaveAs(_excelFilePath);
            }
        }

        // ─── Private helpers ──────────────────────────────────────────────────

        /// <summary>Tìm row SUBTOTAL (nhận dạng bằng formula SUBTOTAL tại cột TIỀN THU).</summary>
        private int FindSubtotalRow(IXLWorksheet worksheet)
        {
            foreach (var row in worksheet.RowsUsed())
            {
                int rn = row.RowNumber();
                if (rn < DATA_START_ROW)
                    continue;
                var cell = row.Cell(COL_TIENTHU);
                if (
                    cell.HasFormula
                    && cell.FormulaA1.StartsWith("SUBTOTAL", StringComparison.OrdinalIgnoreCase)
                )
                    return rn;
            }
            return -1;
        }

        /// <summary>Tìm row có MÃ khớp trong sheet. Trả về -1 nếu không tìm thấy.</summary>
        private int FindRowByMa(IXLWorksheet worksheet, string ma)
        {
            foreach (var row in worksheet.RowsUsed())
            {
                if (row.RowNumber() <= 2)
                    continue;
                if (row.Cell(COL_MA).GetString() == ma)
                    return row.RowNumber();
            }
            return -1;
        }

        /// <summary>
        /// Tìm row cho đơn không có MÃ HĐ (hàng sỉ) bằng cách khớp TÊN KH + NGÀY LẤY.
        /// Chỉ khớp row có MÃ rỗng (tránh overwrite đơn có MÃ trùng tên KH).
        /// Trả về -1 nếu không tìm thấy → sẽ append row mới.
        /// </summary>
        private int FindRowByTenKhNgay(IXLWorksheet worksheet, string tenKh, string ngayLay)
        {
            if (string.IsNullOrWhiteSpace(tenKh))
                return -1;
            // Normalize ngày: chỉ lấy dd-MM (bỏ năm và dấu chấm cuối)
            string normNgay = "";
            if (!string.IsNullOrWhiteSpace(ngayLay))
            {
                var parts = ngayLay.TrimEnd('.').Split('-');
                if (parts.Length >= 2)
                    normNgay = $"{parts[0]}-{parts[1]}";
            }
            foreach (var row in worksheet.RowsUsed())
            {
                if (row.RowNumber() <= 2)
                    continue;
                // Chỉ xét row mà cột MÃ rỗng (hàng sỉ)
                if (!string.IsNullOrWhiteSpace(row.Cell(COL_MA).GetString()))
                    continue;
                string rowTenKh = row.Cell(COL_TENKH).GetString().Trim();
                if (!rowTenKh.Equals(tenKh.Trim(), StringComparison.OrdinalIgnoreCase))
                    continue;
                // Khớp ngày nếu có
                if (!string.IsNullOrEmpty(normNgay))
                {
                    string rowNgay = row.Cell(COL_NGAYLAY).GetString().TrimEnd('.');
                    var rParts = rowNgay.Split('-');
                    string rowNormNgay = rParts.Length >= 2 ? $"{rParts[0]}-{rParts[1]}" : rowNgay;
                    if (!rowNormNgay.Equals(normNgay, StringComparison.OrdinalIgnoreCase))
                        continue;
                }
                return row.RowNumber();
            }
            return -1;
        }

        // ── Mapping tên cột → index (dùng cho UpdateInvoiceFields) ─────────────
        private static readonly Dictionary<string, int> _colNameToIndex = new Dictionary<
            string,
            int
        >(StringComparer.OrdinalIgnoreCase)
        {
            { "TÌNH TRẠNG TT", COL_TINHTRANG },
            { "SHOP", COL_SHOP },
            { "TÊN KH", COL_TENKH },
            { "MÃ", COL_MA },
            { "ĐỊA CHỈ", COL_DIACHI },
            { "QUẬN", COL_QUAN },
            { "TIỀN THU", COL_TIENTHU },
            { "TIỀN SHIP", COL_TIENSHIP },
            { "TIỀN HÀNG", COL_TIENHANG },
            { "NGƯỜI ĐI", COL_NGUOIDI },
            { "NGƯỜI LẤY", COL_NGUOILAY },
            { "NGÀY LẤY", COL_NGAYLAY },
            { "GHI CHÚ", COL_GHICHU },
            { "ỨNG TIỀN", COL_UNGIEN },
            { "HÀNG TỒN", COL_HANGTON },
            { "FAIL", COL_FAIL },
        };

        /// <summary>
        /// Danh sách tên cột có thể chỉnh sửa (dùng để build UI checkbox).
        /// Bỏ qua MÃ (key định danh) và các cột ẩn (COL1-3).
        /// </summary>
        public static IReadOnlyList<string> EditableColumnNames { get; } =
            new[]
            {
                "TÌNH TRẠNG TT",
                "SHOP",
                "TÊN KH",
                "ĐỊA CHỈ",
                "QUẬN",
                "TIỀN THU",
                "TIỀN SHIP",
                "TIỀN HÀNG",
                "NGƯỜI ĐI",
                "NGƯỜI LẤY",
                "NGÀY LẤY",
                "GHI CHÚ",
                "ỨNG TIỀN",
                "HÀNG TỒN",
                "FAIL",
            };

        /// <summary>
        /// Cập nhật các field của invoice theo MÃ HĐ.
        /// Nếu sheetName != null → chỉ tìm trong sheet đó.
        /// Nếu sheetName == null → quét toàn bộ sheets (đơn không rõ ngày).
        /// </summary>
        /// <param name="sheetName">Tên sheet cụ thể, hoặc null để quét tất cả sheets.</param>
        /// <param name="maList">Danh sách mã HĐ cần cập nhật.</param>
        /// <param name="fieldValues">Dict: tên cột → giá trị mới. Key phải là tên trong EditableColumnNames.</param>
        /// <returns>Tuple: (số dòng đã cập nhật, danh sách mã không tìm thấy)</returns>
        public (int updated, List<string> notFound) UpdateInvoiceFields(
            string sheetName,
            IEnumerable<string> maList,
            Dictionary<string, string> fieldValues
        )
        {
            int updated = 0;
            var notFound = new List<string>();

            if (fieldValues == null || fieldValues.Count == 0)
                return (0, notFound);

            using (var workbook = new XLWorkbook(_excelFilePath))
            {
                // Lấy danh sách sheet cần tìm (1 sheet cụ thể hoặc toàn bộ)
                IEnumerable<IXLWorksheet> sheetsToSearch =
                    sheetName != null
                        ? workbook.TryGetWorksheet(sheetName, out var ws)
                            ? new[] { ws }
                            : Array.Empty<IXLWorksheet>()
                        : workbook.Worksheets;

                foreach (var ma in maList)
                {
                    if (string.IsNullOrWhiteSpace(ma))
                        continue;

                    bool found = false;
                    foreach (var worksheet in sheetsToSearch)
                    {
                        int rowNum = FindRowByMa(worksheet, ma.Trim());
                        if (rowNum < 0)
                            continue;

                        found = true;
                        ApplyFieldValuesToRow(worksheet, rowNum, fieldValues);
                        updated++;
                        break; // MÃ unique → dừng sau khi tìm thấy
                    }

                    if (!found)
                        notFound.Add(ma);
                }

                // Nếu sheetName có sẵn nhưng sheet không tồn tại
                if (sheetName != null && !workbook.TryGetWorksheet(sheetName, out _))
                    notFound.AddRange(maList);

                workbook.SaveAs(_excelFilePath);
            }

            return (updated, notFound);
        }

        /// <summary>
        /// Overload không cần chỉ định sheet — quét toàn bộ sheets.
        /// </summary>
        public (int updated, List<string> notFound) UpdateInvoiceFields(
            IEnumerable<string> maList,
            Dictionary<string, string> fieldValues
        ) => UpdateInvoiceFields(null, maList, fieldValues);

        /// <summary>
        /// Ghi fieldValues vào 1 row cụ thể trong worksheet.
        /// Tái sử dụng bởi cả hai overload UpdateInvoiceFields.
        /// </summary>
        private static void ApplyFieldValuesToRow(
            IXLWorksheet worksheet,
            int rowNum,
            Dictionary<string, string> fieldValues
        )
        {
            foreach (var kv in fieldValues)
            {
                if (!_colNameToIndex.TryGetValue(kv.Key, out int colIdx))
                    continue;
                var cell = worksheet.Cell(rowNum, colIdx);
                if (decimal.TryParse(kv.Value, out decimal num))
                    cell.SetValue(num);
                else
                    cell.SetValue(kv.Value);
            }
        }

        /// <summary>
        /// Chuyển đổi TÌNH TRẠNG từ OCR → text ghi vào GHI CHÚ (nếu có nhãn đặc biệt).
        /// Cột TÌNH TRẠNG TT trong Excel luôn = "hàng sỉ".
        /// Các nhãn đặc biệt (KO THU SHIP, THU SHIP, đã CK, v.v.) được chuyển sang GHI CHÚ.
        /// "hàng sỉ" thuần túy (không kèm nhãn đặc biệt) → không cần ghi GHI CHÚ.
        /// </summary>
        private static string BuildGhiChuFromTinhTrang(string tinhTrang, string invoiceType)
        {
            // Ưu tiên: nếu là ship-only → dùng nhãn từ INVOICE_TYPE (chuẩn nhất)
            if (invoiceType == "SHIP_ONLY_FREE")
                return "KO THU SHIP";
            if (invoiceType == "SHIP_ONLY_PAID")
                return "THU SHIP";

            // Với đơn COD: lọc các nhãn đặc biệt ra khỏi TÌNH TRẠNG (bỏ "hàng sỉ")
            if (string.IsNullOrEmpty(tinhTrang))
                return "";

            var parts = tinhTrang
                .Split(new[] { '|', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim())
                .Where(p =>
                    !string.IsNullOrEmpty(p)
                    && !p.Equals("hàng sỉ", StringComparison.OrdinalIgnoreCase)
                    && !p.Equals("hang si", StringComparison.OrdinalIgnoreCase)
                    && !p.Equals("hs", StringComparison.OrdinalIgnoreCase)
                )
                .ToList();

            return string.Join(" | ", parts);
        }

        /// <summary>
        /// Kiểm tra shopVal có trông như ghi chú thủ công không.
        /// Chỉ true nếu SHOP chứa từ khoá ghi chú (dùng AppConstants.SHOP_EXCLUSION_PATTERN).
        /// Tên shop hợp lệ như "ĐOÀN NGÂN CHÂU" sẽ không bị nhận nhầm.
        /// </summary>
        private static bool ShopLooksLikeNote(string shopVal)
        {
            if (string.IsNullOrWhiteSpace(shopVal))
                return false;
            return System.Text.RegularExpressions.Regex.IsMatch(
                shopVal,
                AppConstants.SHOP_EXCLUSION_PATTERN,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
            );
        }

        /// <summary>
        /// Ghi toàn bộ data fields vào một hàng Excel.
        /// isMissing = true → MÃ rỗng, tô nền đỏ đậm để tracking.
        /// data["MISSING_FIELDS"] = "TÊN KH,MÃ,..." → tô đỏ nhạt từng cell bị thiếu.
        /// </summary>
        private void WriteDataRow(
            IXLWorksheet worksheet,
            int targetRow,
            Dictionary<string, string> data,
            string shopVal,
            string ghichuVal,
            string ma,
            bool isFail,
            bool isMissing
        )
        {
            // Map tên field → số cột để tô màu cell thiếu
            var fieldToCol = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                { "SHOP", COL_SHOP },
                { "TÊN KH", COL_TENKH },
                { "MÃ", COL_MA },
                { "ĐỊA CHỈ", COL_DIACHI },
                { "QUẬN", COL_QUAN },
                { "TIỀN THU", COL_TIENTHU },
                { "TIỀN SHIP", COL_TIENSHIP },
                { "NGÀY LẤY", COL_NGAYLAY },
                { "GHI CHÚ", COL_GHICHU },
            };
            var missingSet = new HashSet<string>(
                (data.GetValueOrDefault("MISSING_FIELDS", "")).Split(
                    new[] { ',' },
                    StringSplitOptions.RemoveEmptyEntries
                ),
                StringComparer.OrdinalIgnoreCase
            );

            // ── TÌNH TRẠNG logic (theo Excel mẫu) ─────────────────────────────────
            // - Đơn có MÃ HĐ (COD/SHIP_ONLY) → TÌNH TRẠNG TT = "" (bỏ trống)
            // - Đơn KHÔNG có MÃ (hàng sỉ thực sự) → TÌNH TRẠNG TT = "hàng sỉ"
            // Các nhãn đặc biệt (KO THU SHIP, THU SHIP, đã CK, v.v.) → ghi vào GHI CHÚ.
            string tinhTrangVal = data.GetValueOrDefault("TÌNH TRẠNG", "");
            string invType = data.GetValueOrDefault("INVOICE_TYPE", "COD");

            // Xây GHI CHÚ: lấy các nhãn đặc biệt (loại trừ "hàng sỉ") từ TÌNH TRẠNG
            // rồi merge với ghichuVal gốc (từ caller)
            string tinhTrangNote = BuildGhiChuFromTinhTrang(tinhTrangVal, invType);
            if (!string.IsNullOrEmpty(tinhTrangNote))
            {
                ghichuVal = string.IsNullOrEmpty(ghichuVal)
                    ? tinhTrangNote
                    : tinhTrangNote + " | " + ghichuVal;
            }

            // COD_PLUS_SHIP: đảm bảo GHI CHÚ luôn có note "THU X + SHIP"
            // (safety net: kể cả khi OCR service đã set, hoặc khi load từ log cũ)
            if (invType == "COD_PLUS_SHIP")
            {
                string thuStr0 = data.GetValueOrDefault("TIỀN THU", "");
                string thuNote = $"THU {thuStr0} + SHIP";
                if (string.IsNullOrEmpty(ghichuVal))
                    ghichuVal = thuNote;
                else if (!ghichuVal.Contains("+ SHIP"))
                    ghichuVal = thuNote + " | " + ghichuVal;
            }

            // Đơn có MÃ HĐ → bỏ trống TÌNH TRẠNG TT
            // Đơn KHÔNG có MÃ (bất kể COD hay SHIP_ONLY) → luôn ghi "hàng sỉ"
            // THU SHIP / KO THU SHIP vẫn ghi vào GHI CHÚ như bình thường
            bool isShipOnlyType = invType == "SHIP_ONLY_FREE" || invType == "SHIP_ONLY_PAID";
            string tinhTrangCell = string.IsNullOrWhiteSpace(ma) ? "hàng sỉ" : "";
            worksheet.Cell(targetRow, COL_TINHTRANG).Value = tinhTrangCell;
            worksheet.Cell(targetRow, COL_SHOP).Value = shopVal;
            worksheet.Cell(targetRow, COL_TENKH).Value = data.GetValueOrDefault("TÊN KH", "");
            worksheet.Cell(targetRow, COL_MA).Value = ma;
            worksheet.Cell(targetRow, COL_DIACHI).Value = data.GetValueOrDefault("ĐỊA CHỈ", "");
            worksheet.Cell(targetRow, COL_QUAN).Value = data.GetValueOrDefault("QUẬN", "");

            // ── TIỀN THU + TIỀN SHIP ───────────────────────────────────────────
            // Dùng InvariantCulture để "7.28" luôn parse đúng bất kể locale máy tính.
            //
            // Dữ liệu từ OCR/log:
            //   TIỀN THU  = Tổng thanh toán khách trả (đã bao gồm giảm giá, KHÔNG cộng ship)
            //   TIỀN SHIP = Phí ship thực tế cty trả hộ khách
            //
            // Ghi vào Excel:
            //   COD           : TIỀN THU = thuVal (giữ nguyên)         | TIỀN SHIP = shipVal
            //   COD_PLUS_SHIP : TIỀN THU = thuVal + shipVal (thu cả 2) | TIỀN SHIP = shipVal
            //   SHIP_ONLY_FREE: TIỀN THU = 0                           | TIỀN SHIP = shipVal
            //   SHIP_ONLY_PAID: TIỀN THU = shipVal (thu = ship)        | TIỀN SHIP = shipVal
            string thuStr = data.GetValueOrDefault("TIỀN THU", "0");
            string shipStr = data.GetValueOrDefault("TIỀN SHIP", "0");
            double.TryParse(
                shipStr,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out double shipVal
            );
            double.TryParse(
                thuStr,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out double thuVal
            );

            // Ghi TIỀN THU theo loại đơn
            double tienThuToWrite = invType switch
            {
                "SHIP_ONLY_FREE"  => 0,              // không thu tiền khách
                "SHIP_ONLY_PAID"  => shipVal,         // thu đúng bằng tiền ship
                "COD_PLUS_SHIP"   => thuVal + shipVal, // thu tiền hàng + ship
                _                 => thuVal,           // COD: giữ nguyên Tổng thanh toán từ HĐ
            };
            worksheet.Cell(targetRow, COL_TIENTHU).Value = tienThuToWrite;
            worksheet.Cell(targetRow, COL_TIENTHU).Style.NumberFormat.Format = "#,##0";

            // Ghi TIỀN SHIP
            worksheet.Cell(targetRow, COL_TIENSHIP).Value = shipVal;
            worksheet.Cell(targetRow, COL_TIENSHIP).Style.NumberFormat.Format = "#,##0";

            // TIỀN HÀNG: formula theo loại đơn
            //   COD           : =TIENTHU - TIENSHIP  (= Tổng thanh toán)
            //   COD_PLUS_SHIP : =TIENTHU - TIENSHIP  (= tiền hàng thuần, đã bao gồm ship trong TIENTHU)
            //   SHIP_ONLY_FREE: =-TIENSHIP            (không thu ship → tiền hàng âm)
            //   SHIP_ONLY_PAID: =+TIENSHIP            (thu ship → tiền hàng = ship)
            string thuCol = ColLetter(COL_TIENTHU);
            string shipCol = ColLetter(COL_TIENSHIP);
            var hangCell = worksheet.Cell(targetRow, COL_TIENHANG);
            hangCell.FormulaA1 = invType switch
            {
                "SHIP_ONLY_FREE" => $"-{shipCol}{targetRow}",
                "SHIP_ONLY_PAID" => $"{shipCol}{targetRow}",
                _ => $"{thuCol}{targetRow}-{shipCol}{targetRow}", // COD và COD_PLUS_SHIP đều trừ ship
            };
            hangCell.Style.NumberFormat.Format = "#,##0";

            worksheet.Cell(targetRow, COL_NGUOIDI).Value = data.GetValueOrDefault("NGƯỜI ĐI", "");
            worksheet.Cell(targetRow, COL_NGUOILAY).Value = data.GetValueOrDefault("NGƯỜI LẤY", "");
            worksheet.Cell(targetRow, COL_NGAYLAY).Value = data.GetValueOrDefault("NGÀY LẤY", "");
            worksheet.Cell(targetRow, COL_GHICHU).Value = ghichuVal;
            worksheet.Cell(targetRow, COL_UNGIEN).Value = data.GetValueOrDefault("ỨNG TIỀN", "");
            worksheet.Cell(targetRow, COL_HANGTON).Value = data.GetValueOrDefault("HÀNG TỒN", "");
            worksheet.Cell(targetRow, COL_FAIL).Value = data.GetValueOrDefault("FAIL", "");

            // COL1 = 1 cho TẤT CẢ đơn (kể cả ship-only không có MÃ)
            worksheet.Cell(targetRow, COL_COL1).Value = 1;

            // COL2 = tiền ship (số) khi là SHIP_ONLY_FREE — cột theo dõi ship cty chịu
            if (invType == "SHIP_ONLY_FREE")
                worksheet.Cell(targetRow, COL_COL2).Value = shipVal;
            else
                worksheet.Cell(targetRow, COL_COL2).Value = data.GetValueOrDefault("COL2", "");
            worksheet.Cell(targetRow, COL_COL3).Value = data.GetValueOrDefault("COL3", "");

            // Tô đỏ nhạt từng cell bị thiếu
            foreach (var fieldName in missingSet)
            {
                if (fieldToCol.TryGetValue(fieldName, out int col))
                    worksheet.Cell(targetRow, col).Style.Fill.BackgroundColor = XLColor.FromHtml(
                        AppConstants.COLOR_FAIL_ROW
                    );
            }

            // MÃ rỗng → tô đỏ đậm (ship-only không có MÃ là bình thường → không tô)
            if (isMissing && !isShipOnlyType)
            {
                worksheet.Cell(targetRow, COL_MA).Style.Fill.BackgroundColor = XLColor.FromHtml(
                    AppConstants.COLOR_MISSING_MA
                );
            }
        }

        /// <summary>Tạo bảng tổng kết bên phải (per NGƯỜI ĐI), mỗi người 1 block 7 dòng.</summary>
        private void BuildRightSummary(
            IXLWorksheet worksheet,
            List<string> distinctNguoiDis,
            int startRow,
            string rThu,
            string rShip,
            string rNguoiDi,
            string rShop
        )
        {
            string nguoiLayColL = ColLetter(COL_NGUOILAY);
            string ngayLayColL = ColLetter(COL_NGAYLAY);
            int curRow = startRow;

            foreach (string nd in distinctNguoiDis)
            {
                int b0 = curRow;
                int b1 = curRow + 1;
                int b2 = curRow + 2;
                int b3 = curRow + 3;
                int b4 = curRow + 4;
                int b5 = curRow + 5;
                int b6 = curRow + 6;

                // Header block
                SetBoldYellow(worksheet.Cell(b0, COL_NGUOIDI), nd, XLColor.LightSteelBlue);
                SetBoldYellow(worksheet.Cell(b0, COL_NGUOILAY), "Tiền Thu", XLColor.LightSteelBlue);
                SetBoldYellow(worksheet.Cell(b0, COL_NGAYLAY), "Số đơn", XLColor.LightSteelBlue);

                // Data rows
                // COUNTIFS theo SHOP (không rỗng) thay vì MÃ — để đếm đúng cả đơn hàng si không có MÃ
                worksheet.Cell(b1, COL_NGUOIDI).Value = "TỔNG ĐƠN NHẬN";
                worksheet.Cell(b1, COL_NGUOILAY).FormulaA1 = $"SUMIFS({rThu},{rNguoiDi},\"{nd}\")";
                worksheet.Cell(b1, COL_NGUOILAY).Style.NumberFormat.Format = "#,##0";
                worksheet.Cell(b1, COL_NGAYLAY).FormulaA1 =
                    $"COUNTIFS({rShop},\"<>\",{rNguoiDi},\"{nd}\")";

                worksheet.Cell(b2, COL_NGUOIDI).Value = "tiền ship";
                worksheet.Cell(b2, COL_NGUOILAY).FormulaA1 =
                    $"-SUMIFS({rShip},{rNguoiDi},\"{nd}\")";
                worksheet.Cell(b2, COL_NGUOILAY).Style.NumberFormat.Format = "#,##0";
                // TODO: tiền lấy — sẽ có logic riêng sau, tạm comment
                // worksheet.Cell(b3, COL_NGUOIDI).Value = "tiền lấy";
                // worksheet.Cell(b3, COL_NGUOILAY).FormulaA1 =
                //     $"-{ngayLayColL}{b1}*{(int)AppConstants.PHI_SHIP_MOI_DON}";

                worksheet.Cell(b4, COL_NGUOIDI).Value = "đơn trả";
                worksheet.Cell(b4, COL_NGUOIDI).Style.Font.FontColor = XLColor.Red;

                worksheet.Cell(b5, COL_NGUOIDI).Value = "đơn cũ ck";
                worksheet.Cell(b5, COL_NGUOIDI).Style.Font.FontColor = XLColor.Red;

                // Tổng cuối block
                var cTotal = worksheet.Cell(b6, COL_NGUOILAY);
                cTotal.FormulaA1 = $"SUBTOTAL(9,{nguoiLayColL}{b1}:{nguoiLayColL}{b5})";
                cTotal.Style.Font.Bold = true;
                cTotal.Style.Fill.BackgroundColor = XLColor.LightBlue;
                cTotal.Style.NumberFormat.Format = "#,##0";

                var cTotalDon = worksheet.Cell(b6, COL_NGAYLAY);
                cTotalDon.FormulaA1 = $"{ngayLayColL}{b1}";
                cTotalDon.Style.Font.Bold = true;
                cTotalDon.Style.Fill.BackgroundColor = XLColor.LightBlue;

                curRow += SUMMARY_RIGHT_BLOCK_HEIGHT + 1;
            }
        }

        /// <summary>Tạo bảng tổng kết bên trái (per SHOP), mỗi shop 1 block 8 dòng.</summary>
        private void BuildLeftSummary(
            IXLWorksheet worksheet,
            List<string> distinctShops,
            int startRow,
            string firstNgay,
            string rThu,
            string rShip,
            string rShop,
            string rNgay,
            string rMa,
            string tenkhColL,
            string diachiColL
        )
        {
            int curRow = startRow;

            foreach (string shop in distinctShops)
            {
                int r0 = curRow; // header: tên shop
                int r1 = curRow + 1; // COD (tiền thu + số đơn theo ngày)
                int r2 = curRow + 2; // Trừ tiền ship
                int r3 = curRow + 3; // Đơn trả & CK (đỏ, manual)
                int r4 = curRow + 4; // Tiền hàng HCM (subtotal r1..r3)
                int r5 = curRow + 5; // Đơn đơn (đỏ, manual)
                int r6 = curRow + 6; // Nợ cũ (đỏ, manual)
                int r7 = curRow + 7; // THANH TOÁN

                string aShop = $"$A${r0}";
                // Tên shop ở COL_TINHTRANG (cột A), nhưng SUMIFS/COUNTIFS cần so với data cột B (COL_SHOP).
                // Dùng string literal thay vì cell ref để tránh lệch cột.
                string shopCriteria = $"\"{shop}\"";

                // ── R0: Header — tên shop ─────────────────────────────────────
                SetBoldYellow(worksheet.Cell(r0, COL_TINHTRANG), shop, XLColor.LightSteelBlue);
                SetBoldYellow(worksheet.Cell(r0, COL_SHOP), "", XLColor.LightSteelBlue);
                SetBoldYellow(worksheet.Cell(r0, COL_TENKH), "Tiền", XLColor.LightSteelBlue);
                SetBoldYellow(worksheet.Cell(r0, COL_MA), "Số đơn", XLColor.LightSteelBlue);

                // ── R1: COD — tổng tất cả ngày của shop ──────────────────────
                worksheet.Cell(r1, COL_SHOP).Value = "cod";
                worksheet.Cell(r1, COL_TENKH).FormulaA1 = $"SUMIFS({rThu},{rShop},{shopCriteria})";
                worksheet.Cell(r1, COL_TENKH).Style.NumberFormat.Format = "#,##0";
                // COUNTIFS: đếm đơn có SHOP = tên shop (không rỗng bao gồm cả hàng sỉ không có MÃ)
                worksheet.Cell(r1, COL_MA).FormulaA1 = $"COUNTIFS({rShop},{shopCriteria})";

                // ── R2: Trừ tiền ship ─────────────────────────────────────────
                worksheet.Cell(r2, COL_SHOP).Value = "Trừ Tiền Ship";
                worksheet.Cell(r2, COL_TENKH).FormulaA1 =
                    $"-SUMIFS({rShip},{rShop},{shopCriteria})";
                worksheet.Cell(r2, COL_TENKH).Style.NumberFormat.Format = "#,##0";

                // ── R3: Đơn trả (manual, tô đỏ) ──────────────────────────────
                worksheet.Cell(r3, COL_SHOP).Value = "Đơn trả & c.khoản";
                worksheet.Cell(r3, COL_SHOP).Style.Font.FontColor = XLColor.Red;

                // ── R4: Tiền Hàng HCM (subtotal r1..r3) ──────────────────────
                worksheet.Cell(r4, COL_SHOP).Value = "Tiền Hàng Hcm";
                worksheet.Cell(r4, COL_SHOP).Style.Font.Bold = true;
                worksheet.Cell(r4, COL_TENKH).FormulaA1 =
                    $"SUBTOTAL(9,{tenkhColL}{r1}:{tenkhColL}{r3})";
                worksheet.Cell(r4, COL_TENKH).Style.Font.Bold = true;
                worksheet.Cell(r4, COL_TENKH).Style.NumberFormat.Format = "#,##0";
                worksheet.Cell(r4, COL_MA).FormulaA1 = $"{diachiColL}{r1}";
                worksheet.Cell(r4, COL_MA).Style.Font.Bold = true;

                // ── R5: Đơn đơn (manual, tô đỏ) ─────────────────────────────
                worksheet.Cell(r5, COL_SHOP).Value = "đơn đơn";
                worksheet.Cell(r5, COL_SHOP).Style.Font.FontColor = XLColor.Red;

                // ── R6: Nợ cũ (manual, tô đỏ) ────────────────────────────────
                worksheet.Cell(r6, COL_SHOP).Value = "nợ cũ";
                worksheet.Cell(r6, COL_SHOP).Style.Font.FontColor = XLColor.Red;

                // ── R7: THANH TOÁN ────────────────────────────────────────────
                SetBoldYellow(worksheet.Cell(r7, COL_TINHTRANG), "", XLColor.LightGreen);
                SetBoldYellow(worksheet.Cell(r7, COL_SHOP), "THANH TOÁN", XLColor.LightGreen);
                var cTTVal = worksheet.Cell(r7, COL_TENKH);
                cTTVal.FormulaA1 = $"{tenkhColL}{r4}+{tenkhColL}{r5}+{tenkhColL}{r6}";
                cTTVal.Style.Font.Bold = true;
                cTTVal.Style.Fill.BackgroundColor = XLColor.LightGreen;
                cTTVal.Style.NumberFormat.Format = "#,##0";
                SetBoldYellow(worksheet.Cell(r7, COL_MA), "CK Đủ 100%", XLColor.LightGreen);

                // ── Nền nhạt cho cả block ─────────────────────────────────────
                for (int r = r0; r <= r7; r++)
                for (int c = COL_TINHTRANG; c <= COL_MA; c++)
                {
                    var cell = worksheet.Cell(r, c);
                    var bg = cell.Style.Fill.BackgroundColor;
                    if (bg.Equals(XLColor.NoColor) || bg.Equals(XLColor.White))
                        cell.Style.Fill.BackgroundColor = XLColor.LightCyan;
                }

                // ── Viền ngoài block ──────────────────────────────────────────
                var blockRange = worksheet.Range(r0, COL_TINHTRANG, r7, COL_MA);
                blockRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                blockRange.Style.Border.OutsideBorderColor = XLColor.SteelBlue;
                // Đường kẻ ngang mỏng giữa các dòng
                for (int r = r0; r <= r7; r++)
                    worksheet.Range(r, COL_TINHTRANG, r, COL_MA).Style.Border.BottomBorder =
                        XLBorderStyleValues.Thin;

                curRow += SUMMARY_LEFT_BLOCK_HEIGHT + 1;
            }
        }

        /// <summary>Helper: set Value + Bold + BackgroundColor cho 1 ô.</summary>
        private static void SetBoldYellow(IXLCell cell, string value, XLColor bgColor)
        {
            cell.Value = value;
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = bgColor;
        }

        /// <summary>Helper: set Value + Bold cho 1 ô.</summary>
        private static void SetBold(IXLCell cell, string value)
        {
            cell.Value = value;
            cell.Style.Font.Bold = true;
        }

        /// <summary>Chuyển index cột (1-based) → chữ cột Excel: 1→A, 2→B, ..., 26→Z, 27→AA.</summary>
        private static string ColLetter(int col)
        {
            string result = "";
            while (col > 0)
            {
                col--;
                result = (char)('A' + col % 26) + result;
                col /= 26;
            }
            return result;
        }

        /// <summary>
        /// Upsert invoice: nếu MÃ đã tồn tại trong sheet → ghi đè dòng đó.
        /// Nếu chưa có → thêm dòng mới cuối sheet.
        /// Sheet được chọn theo sheetName (mặc định ngày hôm nay "dd-MM").
        /// </summary>
        public void ExportInvoice(OCRInvoiceData invoice, string sheetName = null)
        {
            if (invoice == null)
                throw new ArgumentNullException(nameof(invoice));

            sheetName ??= DateTime.Now.ToString("dd-MM");

            try
            {
                using (var workbook = new XLWorkbook(_excelFilePath))
                {
                    IXLWorksheet worksheet;
                    if (workbook.TryGetWorksheet(sheetName, out var existingSheet))
                    {
                        worksheet = existingSheet;
                    }
                    else
                    {
                        worksheet = workbook.Worksheets.Add(sheetName);
                        AddHeaderRow(worksheet, DateTime.Now);
                    }

                    // Tìm row có MÃ trùng để ghi đè (upsert)
                    int targetRow = -1;
                    var usedRows = worksheet.RowsUsed();
                    foreach (var row in usedRows)
                    {
                        if (row.RowNumber() <= 2)
                            continue; // bỏ header rows
                        if (row.Cell(COL_MA).GetString() == invoice.SoHoaDon)
                        {
                            targetRow = row.RowNumber();
                            break;
                        }
                    }

                    // Nếu không tìm thấy → append dòng mới
                    if (targetRow < 0)
                    {
                        var lastRow = worksheet.LastRowUsed();
                        targetRow =
                            (lastRow != null && lastRow.RowNumber() >= 3)
                                ? lastRow.RowNumber() + 1
                                : 3;
                    }

                    // Ghi dữ liệu vào targetRow
                    worksheet.Cell(targetRow, COL_SHOP).Value = invoice.Shop ?? "";
                    worksheet.Cell(targetRow, COL_TENKH).Value = invoice.TenKhachHang ?? "";
                    worksheet.Cell(targetRow, COL_MA).Value = invoice.SoHoaDon;
                    worksheet.Cell(targetRow, COL_DIACHI).Value = invoice.DiaChi;
                    worksheet.Cell(targetRow, COL_QUAN).Value = invoice.Quan;
                    worksheet.Cell(targetRow, COL_TIENTHU).Value = invoice.TongThanhToan;
                    worksheet.Cell(targetRow, COL_TIENSHIP).Value = 0;
                    worksheet.Cell(targetRow, COL_TIENHANG).Value = invoice.TongTienHang;
                    worksheet.Cell(targetRow, COL_NGUOIDI).Value = invoice.NguoiDi;
                    worksheet.Cell(targetRow, COL_NGUOILAY).Value = invoice.NguoiLay;
                    worksheet.Cell(targetRow, COL_NGAYLAY).Value = DateTime.Now.ToString(
                        "dd-MM-yyyy."
                    );

                    workbook.SaveAs(_excelFilePath);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error exporting invoice: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Add header rows to new worksheet (row 1 = column headers, row 2 = THU x / NGAY x-x)
        /// </summary>
        private void AddHeaderRow(IXLWorksheet worksheet, DateTime date)
        {
            // Row 1: Column headers (20 columns matching existing sheets)
            var headers = new[]
            {
                "Tình trạng TT",
                "SHOP",
                "TÊN KH",
                "MÃ",
                "ĐỊA CHỈ",
                "QUẬN",
                "TIỀN THU",
                "TIỀN SHIP",
                "TIỀN HÀNG",
                "NGƯỜI ĐI",
                "NGƯỜI LẤY",
                "NGÀY LẤY",
                "GHI CHÚ",
                "ỨNG TIỀN",
                "HÀNG TỒN",
                "FAIL",
                "Column1",
                "Column2",
                "Column3",
            };
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.LightGray;
            }

            // AutoFilter on header row (sort/filter dropdown trên tất cả cột)
            worksheet.Range(1, 1, 1, headers.Length).SetAutoFilter();

            // Row 2: THU x | NGAY x-x (matches existing sheet pattern)
            // Day of week: Mon=THU 2, Tue=THU 3, ... Sun=CHU NHAT
            string thuText;
            if (date.DayOfWeek == DayOfWeek.Sunday)
                thuText = "CHU NHAT";
            else
                thuText = "THU " + ((int)date.DayOfWeek + 1).ToString();

            string ngayText = "NGAY " + date.Day + "-" + date.Month;

            var cellThu = worksheet.Cell(2, COL_SHOP);
            cellThu.Value = thuText;
            cellThu.Style.Font.Bold = true;

            var cellNgay = worksheet.Cell(2, COL_TENKH);
            cellNgay.Value = ngayText;
            cellNgay.Style.Font.Bold = true;
        }
    }
}

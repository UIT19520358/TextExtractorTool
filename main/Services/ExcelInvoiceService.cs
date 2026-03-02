using System;
using System.Collections.Generic;
using System.IO;
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

                // Tìm SUBTOTAL row (nếu đã có) để chèn data trước nó
                int existingSubtotalRow = FindSubtotalRow(worksheet);

                // Next empty data row = trước SUBTOTAL row (nếu có), hoặc sau last used row
                int nextRow = DATA_START_ROW;
                if (existingSubtotalRow > 0)
                {
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
                    bool hasMissingFields = !string.IsNullOrEmpty(data.GetValueOrDefault("MISSING_FIELDS", ""));

                    // Upsert: tìm row có MÃ trùng → ghi đè; nếu MÃ rỗng hoặc không tìm thấy → thêm mới
                    int targetRow = isMissing ? -1 : FindRowByMa(worksheet, ma);
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
                int lastDataRow = DATA_START_ROW - 1;
                foreach (var row in worksheet.RowsUsed())
                {
                    int rn = row.RowNumber();
                    if (rn < DATA_START_ROW)
                        continue;
                    string shopVal = row.Cell(COL_SHOP).GetString();
                    string thuVal = row.Cell(COL_TIENTHU).GetString();
                    if (string.IsNullOrWhiteSpace(shopVal) && string.IsNullOrWhiteSpace(thuVal))
                        continue;
                    lastDataRow = rn;
                }
                if (lastDataRow < DATA_START_ROW)
                    throw new InvalidOperationException("Sheet không có dữ liệu.");

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
                        hangCell.FormulaA1 = $"{thuColLetter}{r}+{shipColLetter}{r}"; // COD: thu + ship
                }

                // ── SUBTOTAL row ───────────────────────────────────────────────
                int subtotalRow = lastDataRow + 2;
                worksheet.Row(lastDataRow + 1).Clear();

                int[] subtotalCols =
                {
                    COL_TIENTHU,
                    COL_TIENSHIP,
                    COL_TIENHANG,
                    COL_NGUOIDI,
                    COL_NGUOILAY,
                    COL_NGAYLAY,
                    COL_GHICHU,
                    COL_UNGIEN,
                    COL_HANGTON,
                    COL_FAIL,
                    COL_COL1,
                    COL_COL2,
                    COL_COL3,
                };
                foreach (int col in subtotalCols)
                {
                    string colLetter = ColLetter(col);
                    var stCell = worksheet.Cell(subtotalRow, col);
                    stCell.FormulaA1 =
                        $"SUBTOTAL(9,{colLetter}{DATA_START_ROW}:{colLetter}{lastDataRow})";
                    stCell.Style.Font.Bold = true;
                    stCell.Style.Fill.BackgroundColor = XLColor.LightYellow;
                }

                // COUNTA cột MÃ để đếm số đơn (SUBTOTAL(9) sai vì FAIL="" = 0)
                var maCountCell = worksheet.Cell(subtotalRow, COL_MA);
                maCountCell.FormulaA1 =
                    $"COUNTA({ColLetter(COL_MA)}{DATA_START_ROW}:{ColLetter(COL_MA)}{lastDataRow})";
                maCountCell.Style.Font.Bold = true;
                maCountCell.Style.Fill.BackgroundColor = XLColor.LightYellow;

                // ── Bảng tổng kết ─────────────────────────────────────────────
                int summaryRow = subtotalRow + 2;

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
                    rMa
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
                { "SHOP",      COL_SHOP    },
                { "TÊN KH",   COL_TENKH   },
                { "MÃ",       COL_MA      },
                { "ĐỊA CHỈ",  COL_DIACHI  },
                { "QUẬN",     COL_QUAN    },
                { "TIỀN THU", COL_TIENTHU },
                { "TIỀN SHIP",COL_TIENSHIP},
                { "NGÀY LẤY", COL_NGAYLAY },
                { "GHI CHÚ",  COL_GHICHU  },
            };
            var missingSet = new HashSet<string>(
                (data.GetValueOrDefault("MISSING_FIELDS", ""))
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries),
                StringComparer.OrdinalIgnoreCase
            );

            worksheet.Cell(targetRow, COL_TINHTRANG).Value = "";
            worksheet.Cell(targetRow, COL_SHOP).Value = shopVal;
            worksheet.Cell(targetRow, COL_TENKH).Value = data.GetValueOrDefault("TÊN KH", "");
            worksheet.Cell(targetRow, COL_MA).Value = ma;
            worksheet.Cell(targetRow, COL_DIACHI).Value = data.GetValueOrDefault("ĐỊA CHỈ", "");
            worksheet.Cell(targetRow, COL_QUAN).Value = data.GetValueOrDefault("QUẬN", "");

            // TIỀN THU + TIỀN SHIP: ưu tiên ghi số, fallback ghi text
            // Dùng InvariantCulture để "7.28" luôn parse thành 7.28 bất kể locale máy tính
            string thuStr = data.GetValueOrDefault("TIỀN THU", "0");
            string shipStr = data.GetValueOrDefault("TIỀN SHIP", "0");
            if (
                double.TryParse(
                    thuStr,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out double thuVal
                )
            )
                worksheet.Cell(targetRow, COL_TIENTHU).Value = thuVal;
            else
                worksheet.Cell(targetRow, COL_TIENTHU).Value = thuStr;
            if (
                double.TryParse(
                    shipStr,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out double shipVal
                )
            )
                worksheet.Cell(targetRow, COL_TIENSHIP).Value = shipVal;
            else
                worksheet.Cell(targetRow, COL_TIENSHIP).Value = shipStr;

            // TIỀN HÀNG: formula theo loại đơn
            //   COD           : =TIENTHU + TIENSHIP  (thu x + ship → hàng = x + ship)
            //   SHIP_ONLY_FREE: =-TIENSHIP            (không thu ship → hàng âm)
            //   SHIP_ONLY_PAID: =+TIENSHIP            (thu ship → hàng = ship)
            string thuCol = ColLetter(COL_TIENTHU);
            string shipCol = ColLetter(COL_TIENSHIP);
            string invType = data.GetValueOrDefault("INVOICE_TYPE", "COD");
            worksheet.Cell(targetRow, COL_TIENHANG).FormulaA1 = invType switch
            {
                "SHIP_ONLY_FREE" => $"-{shipCol}{targetRow}",
                "SHIP_ONLY_PAID" => $"{shipCol}{targetRow}",
                _ => $"{thuCol}{targetRow}+{shipCol}{targetRow}", // COD: thu + ship
            };

            worksheet.Cell(targetRow, COL_NGUOIDI).Value = data.GetValueOrDefault("NGƯỜI ĐI", "");
            worksheet.Cell(targetRow, COL_NGUOILAY).Value = data.GetValueOrDefault("NGƯỜI LẤY", "");
            worksheet.Cell(targetRow, COL_NGAYLAY).Value = data.GetValueOrDefault("NGÀY LẤY", "");
            worksheet.Cell(targetRow, COL_GHICHU).Value = ghichuVal;
            worksheet.Cell(targetRow, COL_UNGIEN).Value = data.GetValueOrDefault("ỨNG TIỀN", "");
            worksheet.Cell(targetRow, COL_HANGTON).Value = data.GetValueOrDefault("HÀNG TỒN", "");
            worksheet.Cell(targetRow, COL_FAIL).Value = data.GetValueOrDefault("FAIL", "");
            worksheet.Cell(targetRow, COL_COL1).Value = data.GetValueOrDefault("COL1", "");
            worksheet.Cell(targetRow, COL_COL2).Value = data.GetValueOrDefault("COL2", "");
            worksheet.Cell(targetRow, COL_COL3).Value = data.GetValueOrDefault("COL3", "");

            // Tô đỏ nhạt từng cell bị thiếu (thay vì tô cả hàng)
            foreach (var fieldName in missingSet)
            {
                if (fieldToCol.TryGetValue(fieldName, out int col))
                    worksheet.Cell(targetRow, col).Style.Fill.BackgroundColor = XLColor.FromHtml(
                        AppConstants.COLOR_FAIL_ROW
                    );
            }

            // MÃ rỗng → tô đỏ đậm hơn (luôn áp dụng kể cả khi có trong missingSet)
            if (isMissing)
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
            string rMa
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
                worksheet.Cell(b1, COL_NGUOIDI).Value = "TỔNG ĐƠN NHẬN";
                worksheet.Cell(b1, COL_NGUOILAY).FormulaA1 = $"SUMIFS({rThu},{rNguoiDi},\"{nd}\")";
                worksheet.Cell(b1, COL_NGAYLAY).FormulaA1 =
                    $"COUNTIFS({rMa},\"<>\",{rNguoiDi},\"{nd}\")";

                worksheet.Cell(b2, COL_NGUOIDI).Value = "tiền ship";
                worksheet.Cell(b2, COL_NGUOILAY).FormulaA1 =
                    $"-SUMIFS({rShip},{rNguoiDi},\"{nd}\")";

                worksheet.Cell(b3, COL_NGUOIDI).Value = "tiền lấy";
                worksheet.Cell(b3, COL_NGUOILAY).FormulaA1 =
                    $"-{ngayLayColL}{b1}*{(int)AppConstants.PHI_SHIP_MOI_DON}";

                worksheet.Cell(b4, COL_NGUOIDI).Value = "đơn trả";
                worksheet.Cell(b4, COL_NGUOIDI).Style.Font.FontColor = XLColor.Red;

                worksheet.Cell(b5, COL_NGUOIDI).Value = "đơn cũ ck";
                worksheet.Cell(b5, COL_NGUOIDI).Style.Font.FontColor = XLColor.Red;

                // Tổng cuối block
                var cTotal = worksheet.Cell(b6, COL_NGUOILAY);
                cTotal.FormulaA1 = $"SUBTOTAL(9,{nguoiLayColL}{b1}:{nguoiLayColL}{b5})";
                cTotal.Style.Font.Bold = true;
                cTotal.Style.Fill.BackgroundColor = XLColor.LightBlue;

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
                int r0 = curRow; // tên shop + header
                int r1 = curRow + 1; // ngày + cod row
                int r2 = curRow + 2; // trừ tiền ship
                int r3 = curRow + 3; // đơn trả & ck
                int r4 = curRow + 4; // tiền hàng HCM (subtotal)
                int r5 = curRow + 5; // đơn đơn (manual)
                int r6 = curRow + 6; // nợ cũ (manual)
                int r7 = curRow + 7; // THANH TOÁN

                string aShop = $"$A${r0}";
                string aNgay = $"$A${r1}";

                // Header: tên shop
                var cShop = worksheet.Cell(r0, COL_TINHTRANG);
                cShop.Value = shop;
                cShop.Style.Font.Bold = true;
                cShop.Style.Fill.BackgroundColor = XLColor.LightYellow;
                SetBold(worksheet.Cell(r0, COL_TENKH), "Tiền");
                SetBold(worksheet.Cell(r0, COL_DIACHI), "Số đơn");

                // Ngày + COD row
                worksheet.Cell(r1, COL_TINHTRANG).Value = firstNgay;
                worksheet.Cell(r1, COL_SHOP).Value = "cod";
                worksheet.Cell(r1, COL_TENKH).FormulaA1 =
                    $"SUMIFS({rThu},{rShop},{aShop},{rNgay},{aNgay})";
                worksheet.Cell(r1, COL_DIACHI).FormulaA1 =
                    $"COUNTIFS({rMa},\"<>\",{rShop},{aShop},{rNgay},{aNgay})";

                // Trừ tiền ship
                worksheet.Cell(r2, COL_SHOP).Value = "Trừ Tiền Ship";
                worksheet.Cell(r2, COL_TENKH).FormulaA1 =
                    $"-SUMIFS({rShip},{rShop},{aShop},{rNgay},{aNgay})";

                // Đơn trả (manual, tô đỏ)
                worksheet.Cell(r3, COL_SHOP).Value = "Đơn trả & c.khoản";
                worksheet.Cell(r3, COL_SHOP).Style.Font.FontColor = XLColor.Red;

                // Tiền Hàng HCM (subtotal r1..r3)
                worksheet.Cell(r4, COL_SHOP).Value = "Tiền Hàng Hcm";
                worksheet.Cell(r4, COL_TENKH).FormulaA1 =
                    $"SUBTOTAL(9,{tenkhColL}{r1}:{tenkhColL}{r3})";
                worksheet.Cell(r4, COL_DIACHI).FormulaA1 = $"{diachiColL}{r1}";

                // Đơn đơn + nợ cũ (manual, tô đỏ)
                worksheet.Cell(r5, COL_SHOP).Value = "đơn đơn";
                worksheet.Cell(r5, COL_SHOP).Style.Font.FontColor = XLColor.Red;
                worksheet.Cell(r6, COL_SHOP).Value = "nợ cũ";
                worksheet.Cell(r6, COL_SHOP).Style.Font.FontColor = XLColor.Red;

                // THANH TOÁN
                var cTT = worksheet.Cell(r7, COL_SHOP);
                cTT.Value = "THANH TOÁN";
                cTT.Style.Font.Bold = true;
                cTT.Style.Fill.BackgroundColor = XLColor.LightGreen;
                var cTTVal = worksheet.Cell(r7, COL_TENKH);
                cTTVal.FormulaA1 = $"{tenkhColL}{r4}+{tenkhColL}{r5}+{tenkhColL}{r6}";
                cTTVal.Style.Font.Bold = true;
                cTTVal.Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(r7, COL_DIACHI).Value = "CK Đủ 100%";

                // Nền LightCyan cho toàn block (trừ ô đã có màu riêng)
                for (int r = r0; r <= r7; r++)
                for (int c = COL_TINHTRANG; c <= COL_QUAN; c++)
                {
                    var bg = worksheet.Cell(r, c).Style.Fill.BackgroundColor;
                    if (bg.Equals(XLColor.NoColor) || bg.Equals(XLColor.White))
                        worksheet.Cell(r, c).Style.Fill.BackgroundColor = XLColor.LightCyan;
                }

                // Restore màu ưu tiên cho header + footer
                cShop.Style.Fill.BackgroundColor = XLColor.LightYellow;
                cTT.Style.Fill.BackgroundColor = XLColor.LightGreen;
                cTTVal.Style.Fill.BackgroundColor = XLColor.LightGreen;

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

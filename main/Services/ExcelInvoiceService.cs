using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace TextInputter.Services
{
    /// <summary>
    /// DTO chứa số liệu tổng hợp per NGƯỜI ĐI đã tính trên UI.
    /// Truyền vào ApplyFormulasAndSummary() để Excel ghi đúng y chang UI.
    /// </summary>
    public class NguoiDiSummary
    {
        public decimal TienThu { get; set; } // TỔNG ĐƠN NHẬN — Tiền Thu
        public decimal SoDon { get; set; } // TỔNG ĐƠN NHẬN — Số đơn
        public decimal TienShipTru { get; set; } // tiền ship (số âm hoặc 0)
        public decimal SoDonGiao { get; set; } // số đơn giao thực tế (SoDon - SoDonGop)
        public int SoDonGop { get; set; } // số đơn gộp
        public decimal TienLay { get; set; } // tiền lấy (số âm hoặc 0)
        public decimal DonLayThucTe { get; set; } // SoDon - SoDonTra - SoDonGop
        public decimal TienDonTra { get; set; } // tổng tiền đơn trả (số âm)
        public int SoDonTra { get; set; } // số đơn trả
        public bool IsAnTam { get; set; } // true = An Tâm → skip ship/lấy/trả
    }

    /// <summary>
    /// DTO chứa số liệu bảng TỔNG HỢP (bên trái) đã tính trên UI.
    /// Truyền vào ApplyFormulasAndSummary() để Excel ghi y chang UI.
    /// </summary>
    public class TongHopSummary
    {
        public decimal TienHang { get; set; } // = TongTienThu (SUM cột TIỀN THU)
        public decimal TruShip { get; set; } // = -TongTienShip (số âm)
        public decimal SoDon { get; set; } // Tổng số đơn
        public decimal TienDonTra { get; set; } // Tổng tiền đơn trả (số âm)
        public int SoDonTra { get; set; } // Số đơn trả
        public decimal TongNegative { get; set; } // Tổng row âm (đơn cũ ck)
        public int SoRowNegative { get; set; } // Số dòng âm
        public string NegativeLabel { get; set; } // Nhãn cho row âm (nếu chỉ 1 dòng)
    }

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

        // Số dòng mỗi block bảng tổng kết bên trái (per SHOP) — bảng đối soát gửi shop.
        private const int SUMMARY_LEFT_BLOCK_HEIGHT = 10;

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

                // Nếu không có SUBTOTAL → xóa các hàng data trống hoàn toàn (toàn 0, không có MÃ/SHOP)
                // Đây là tàn dư của lần export bị bug cũ. Xóa để tránh append sau hàng 0.
                if (existingSubtotalRow < 0)
                {
                    var rowsToDelete = new List<int>();
                    var lastUsedForClean = worksheet.LastRowUsed();
                    int lastClean = lastUsedForClean?.RowNumber() ?? (DATA_START_ROW - 1);
                    for (int r = lastClean; r >= DATA_START_ROW; r--)
                    {
                        string shopC = worksheet.Cell(r, COL_SHOP).GetString().Trim();
                        string maC = worksheet.Cell(r, COL_MA).GetString().Trim();
                        string thuC = worksheet.Cell(r, COL_TIENTHU).GetString().Trim();
                        // Hàng rỗng thật sự: không có SHOP, MÃ, và TIỀN THU rỗng hoặc = "0"
                        bool isEmpty =
                            string.IsNullOrWhiteSpace(shopC)
                            && string.IsNullOrWhiteSpace(maC)
                            && (string.IsNullOrWhiteSpace(thuC) || thuC == "0");
                        if (isEmpty)
                            rowsToDelete.Add(r);
                        else
                            break; // dừng khi gặp hàng có data thật (xóa từ cuối lên)
                    }
                    foreach (int r in rowsToDelete)
                        worksheet.Row(r).Delete();
                }

                // nextRow = vị trí để thêm đơn mới (insert trước SUBTOTAL nếu có)
                int nextRow = DATA_START_ROW;
                if (existingSubtotalRow > 0)
                {
                    // nextRow bắt đầu ngay trước SUBTOTAL; mỗi lần thêm mới sẽ InsertRowAbove 1 row
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

                    // Upsert: tìm row có MÃ trùng → ghi đè; nếu MÃ rỗng hoặc không tìm thấy → thêm mới
                    int targetRow = isMissing ? -1 : FindRowByMa(worksheet, ma);
                    bool isUpdate = targetRow > 0;
                    if (!isUpdate)
                    {
                        // Nếu có SUBTOTAL → chèn 1 row trước SUBTOTAL tại nextRow,
                        // clear nội dung row mới (ClosedXML copy format/value từ row trên),
                        // rồi tăng nextRow để lần sau chèn tiếp ở row kế tiếp.
                        if (existingSubtotalRow > 0)
                        {
                            worksheet.Row(nextRow).InsertRowsAbove(1);
                            worksheet.Row(nextRow).Clear(); // xóa formula/value bị copy từ SUBTOTAL
                            targetRow = nextRow;
                            nextRow++; // lần sau chèn ở row tiếp theo (SUBTOTAL đã dịch xuống)
                        }
                        else
                        {
                            targetRow = nextRow;
                            nextRow++;
                        }
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
        /// <param name="nguoiDiDetails">Số liệu per NGƯỜI ĐI đã tính trên UI — nếu null thì dùng formula thuần.</param>
        /// <param name="tongHop">Số liệu bảng TỔNG HỢP đã tính trên UI — nếu null thì dùng formula thuần.</param>
        public void ApplyFormulasAndSummary(
            string sheetName,
            DateTime sheetDate,
            Dictionary<string, NguoiDiSummary> nguoiDiDetails = null,
            TongHopSummary tongHop = null
        )
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
                // Đồng thời fix TÌNH TRẠNG (no MÃ = "hàng sỉ") cho trường hợp
                // WriteSheetToWorkbook chỉ copy raw values từ dgvInvoice
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

                    // Nếu cell đã có formula SHIP_ONLY (từ source Excel) → giữ nguyên
                    string existingFormula = hangCell.HasFormula ? hangCell.FormulaA1.Trim() : "";
                    bool isShipOnlyFormula =
                        existingFormula == $"-{shipColLetter}{r}"
                        || existingFormula == $"{shipColLetter}{r}";
                    if (isShipOnlyFormula)
                    {
                        // Chỉ fix TÌNH TRẠNG
                        string maCheck = worksheet.Cell(r, COL_MA).GetString().Trim();
                        worksheet.Cell(r, COL_TINHTRANG).Value = string.IsNullOrEmpty(maCheck)
                            ? "hàng sỉ"
                            : "";
                        continue;
                    }

                    // Detect SHIP_ONLY khi cell KHÔNG có formula (sau WriteSheetToWorkbook ghi raw)
                    // SHIP_ONLY: TIỀN THU = 0 trong Excel
                    double thuValue = 0;
                    var thuCell = worksheet.Cell(r, COL_TIENTHU);
                    if (!thuCell.HasFormula)
                        double.TryParse(thuCell.GetString(), out thuValue);
                    else
                    {
                        try
                        {
                            thuValue = thuCell.GetDouble();
                        }
                        catch
                        {
                            thuValue = 0;
                        }
                    }

                    if (thuValue == 0)
                    {
                        // SHIP_ONLY: dùng TIỀN HÀNG raw value để phân biệt FREE vs PAID
                        // FREE: raw = -(ship) → formula =-H
                        // PAID: raw = +(ship) → formula =+H
                        double hangRaw = 0;
                        if (!hangCell.HasFormula)
                            double.TryParse(hangCell.GetString(), out hangRaw);
                        if (hangRaw < 0)
                            hangCell.FormulaA1 = $"-{shipColLetter}{r}"; // SHIP_ONLY_FREE
                        else
                            hangCell.FormulaA1 = $"{shipColLetter}{r}"; // SHIP_ONLY_PAID
                    }
                    else
                    {
                        // COD: TIỀN HÀNG = TIỀN THU - TIỀN SHIP
                        hangCell.FormulaA1 = $"{thuColLetter}{r}-{shipColLetter}{r}";
                    }

                    // TÌNH TRẠNG: no MÃ = "hàng sỉ" (đảm bảo đúng dù dgvInvoice có data cũ)
                    string maVal = worksheet.Cell(r, COL_MA).GetString().Trim();
                    worksheet.Cell(r, COL_TINHTRANG).Value = string.IsNullOrEmpty(maVal)
                        ? "hàng sỉ"
                        : "";

                    // NGÀY LẤY: trim trailing dots (Gemini đôi khi trả "27-02-2026." → gây sai SUMIFS)
                    string ngayVal = worksheet.Cell(r, COL_NGAYLAY).GetString().TrimEnd('.');
                    if (ngayVal != worksheet.Cell(r, COL_NGAYLAY).GetString())
                        worksheet.Cell(r, COL_NGAYLAY).Value = ngayVal;
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
                // Collect distinct (shop, ngày) pairs for left summary
                var distinctShopDates = new List<(string Shop, string Ngay)>();
                for (int r = DATA_START_ROW; r <= lastDataRow; r++)
                {
                    string shop = worksheet.Cell(r, COL_SHOP).GetString().Trim();
                    string nguoiDi = worksheet.Cell(r, COL_NGUOIDI).GetString().Trim();
                    string ngay = worksheet.Cell(r, COL_NGAYLAY).GetString().Trim();
                    if (!string.IsNullOrWhiteSpace(shop) && !distinctShops.Contains(shop))
                        distinctShops.Add(shop);
                    if (!string.IsNullOrWhiteSpace(nguoiDi) && !distinctNguoiDis.Contains(nguoiDi))
                        distinctNguoiDis.Add(nguoiDi);
                    if (!string.IsNullOrWhiteSpace(shop) && !string.IsNullOrWhiteSpace(ngay))
                    {
                        var pair = (shop, ngay);
                        if (
                            !distinctShopDates.Any(p =>
                                p.Shop.Equals(shop, StringComparison.OrdinalIgnoreCase)
                                && p.Ngay.Equals(ngay, StringComparison.OrdinalIgnoreCase)
                            )
                        )
                            distinctShopDates.Add(pair);
                    }
                }
                if (distinctShops.Count == 0)
                    distinctShops.Add(AppConstants.SHOP_DEFAULT);
                if (distinctShopDates.Count == 0)
                    distinctShopDates.Add(
                        (distinctShops[0], sheetDate.ToString(AppConstants.DATE_FORMAT_EXCEL))
                    );

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
                string failColL = ColLetter(COL_FAIL);
                string rFail = $"{failColL}${DATA_START_ROW}:{failColL}${lastDataRow}";

                // ── BẢNG PHẢI: per NGƯỜI ĐI ───────────────────────────────────
                BuildRightSummary(
                    worksheet,
                    distinctNguoiDis,
                    summaryRow,
                    rThu,
                    rShip,
                    rNguoiDi,
                    rMa,
                    rShop,
                    nguoiDiDetails
                );

                // ── BẢNG TRÁI: per SHOP × NGÀY ──────────────────────────────
                BuildLeftSummary(
                    worksheet,
                    distinctShopDates,
                    summaryRow,
                    rThu,
                    rShip,
                    rShop,
                    rNgay,
                    rMa,
                    tenkhColL,
                    diachiColL,
                    tongHop
                );

                // ── AutoFilter trên header row (giống file gốc) ──────────────
                // Header ở row 2, data từ row 3 đến lastDataRow, cột A..P (COL_FAIL)
                int headerRow = DATA_START_ROW - 1; // row 2
                worksheet.Range(headerRow, 1, lastDataRow, COL_FAIL).SetAutoFilter();

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

            // TÌNH TRẠNG — tính lúc xuất Excel (KHÔNG lấy từ OCR log).
            // Rule theo file gốc NGAY 27-2.xlsx:
            //   - Đơn không có MÃ HĐ → "hàng sỉ" (bất kể INVOICE_TYPE: COD hay SHIP_ONLY)
            //   - Đơn có MÃ → để trống
            {
                string maHD = (ma ?? "").Trim();
                string tinhTrang = string.IsNullOrEmpty(maHD) ? "hàng sỉ" : "";
                worksheet.Cell(targetRow, COL_TINHTRANG).Value = tinhTrang;
            }
            worksheet.Cell(targetRow, COL_SHOP).Value = shopVal;
            worksheet.Cell(targetRow, COL_TENKH).Value = data.GetValueOrDefault("TÊN KH", "");
            worksheet.Cell(targetRow, COL_MA).Value = ma;
            worksheet.Cell(targetRow, COL_DIACHI).Value = data.GetValueOrDefault("ĐỊA CHỈ", "");
            worksheet.Cell(targetRow, COL_QUAN).Value = data.GetValueOrDefault("QUẬN", "");

            // TIỀN THU + TIỀN SHIP: ưu tiên ghi số, fallback ghi text
            // Dùng InvariantCulture để "7.28" luôn parse thành 7.28 bất kể locale máy tính
            //
            // Convention (giống file mẫu NGAY 27-2.xlsx):
            //   TIỀN THU  = tổng tiền thu từ khách (bao gồm cả ship)
            //   TIỀN HÀNG = TIỀN THU - TIỀN SHIP (tiền hàng thuần)
            //
            // OCR trích TIỀN THU = tiền hàng thuần (chưa có ship) cho COD,
            // nên cần cộng thêm TIỀN SHIP vào TIỀN THU khi ghi Excel.
            string thuStr = data.GetValueOrDefault("TIỀN THU", "0");
            string shipStr = data.GetValueOrDefault("TIỀN SHIP", "0");
            string invType = data.GetValueOrDefault("INVOICE_TYPE", "COD");

            double.TryParse(
                thuStr,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out double thuVal
            );
            double.TryParse(
                shipStr,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out double shipVal
            );

            // COD: TIỀN THU (Excel) = tiền hàng (OCR) + tiền ship = tổng thu từ khách
            if (invType == "COD" && thuVal > 0)
                thuVal += shipVal;

            worksheet.Cell(targetRow, COL_TIENTHU).Value = thuVal;
            worksheet.Cell(targetRow, COL_TIENSHIP).Value = shipVal;

            // TIỀN HÀNG: formula theo loại đơn
            //   COD           : =TIENTHU - TIENSHIP  (tổng thu - ship = tiền hàng thuần)
            //   SHIP_ONLY_FREE: =-TIENSHIP            (không thu ship → hàng âm)
            //   SHIP_ONLY_PAID: =+TIENSHIP            (thu ship → hàng = ship)
            string thuCol = ColLetter(COL_TIENTHU);
            string shipCol = ColLetter(COL_TIENSHIP);
            worksheet.Cell(targetRow, COL_TIENHANG).FormulaA1 = invType switch
            {
                "SHIP_ONLY_FREE" => $"-{shipCol}{targetRow}",
                "SHIP_ONLY_PAID" => $"{shipCol}{targetRow}",
                _ => $"{thuCol}{targetRow}-{shipCol}{targetRow}", // COD: tổng thu - ship = hàng
            };

            worksheet.Cell(targetRow, COL_NGUOIDI).Value = data.GetValueOrDefault("NGƯỜI ĐI", "");
            worksheet.Cell(targetRow, COL_NGUOILAY).Value = data.GetValueOrDefault("NGƯỜI LẤY", "");
            // NGÀY LẤY: trim trailing dots/periods — Gemini đôi khi trả "27-02-2026." thay vì "27-02-2026"
            // Nếu không trim → SUMIFS/COUNTIFS sẽ không match được giữa "27-02-2026." và "27-02-2026"
            worksheet.Cell(targetRow, COL_NGAYLAY).Value = data.GetValueOrDefault("NGÀY LẤY", "")
                .TrimEnd('.');
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
            string rMa,
            string rShop,
            Dictionary<string, NguoiDiSummary> nguoiDiDetails = null
        )
        {
            string nguoiLayColL = ColLetter(COL_NGUOILAY);
            string ngayLayColL = ColLetter(COL_NGAYLAY);
            string failColL = ColLetter(COL_FAIL);

            // Ranges for đơn gộp & đơn trả detection (TÊN KH + ĐỊA CHỈ + FAIL)
            string rFail =
                $"{failColL}${DATA_START_ROW}:{failColL}${worksheet.LastRowUsed()?.RowNumber() ?? DATA_START_ROW}";

            int curRow = startRow;

            foreach (string nd in distinctNguoiDis)
            {
                bool isAnTam = nd.Equals(
                    AppConstants.NGUOI_DI_DEFAULT,
                    StringComparison.OrdinalIgnoreCase
                );

                // Lấy số liệu đã tính từ UI (nếu có)
                NguoiDiSummary detail = null;
                nguoiDiDetails?.TryGetValue(nd, out detail);

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

                // ── TỔNG ĐƠN NHẬN ────────────────────────────────────────────
                worksheet.Cell(b1, COL_NGUOIDI).Value = "TỔNG ĐƠN NHẬN";
                worksheet.Cell(b1, COL_NGUOILAY).FormulaA1 = $"SUMIFS({rThu},{rNguoiDi},\"{nd}\")";
                worksheet.Cell(b1, COL_NGAYLAY).FormulaA1 =
                    $"COUNTIFS({rShop},\"<>\",{rNguoiDi},\"{nd}\")";

                if (!isAnTam)
                {
                    // ── tiền ship ─────────────────────────────────────────────
                    // UI: -(TienShip - SoDonGiao × 5k) trong đó SoDonGiao = SoDon - SoDonGop
                    // Dùng giá trị từ UI khi có đơn gộp (formula thuần không detect được đơn gộp)
                    worksheet.Cell(b2, COL_NGUOIDI).Value = "tiền ship";
                    if (detail != null && detail.SoDonGop > 0)
                    {
                        // Có đơn gộp → ghi giá trị đã tính từ UI
                        worksheet.Cell(b2, COL_NGUOILAY).Value = (double)detail.TienShipTru;
                        worksheet.Cell(b2, COL_NGAYLAY).Value = $"giao {detail.SoDonGiao:N0}";
                    }
                    else
                    {
                        // Không có đơn gộp → formula giống UI
                        worksheet.Cell(b2, COL_NGUOILAY).FormulaA1 =
                            $"-SUMIFS({rShip},{rNguoiDi},\"{nd}\")+COUNTIFS({rShop},\"<>\",{rNguoiDi},\"{nd}\")*{AppConstants.PHI_SHIP_MOI_DON}";
                        if (detail != null)
                            worksheet.Cell(b2, COL_NGAYLAY).Value = $"giao {detail.SoDonGiao:N0}";
                    }

                    // ── tiền lấy ──────────────────────────────────────────────
                    // UI: -((SoDon - SoDonTra - SoDonGop) × 2k)
                    // Formula thuần chỉ trừ SoDonTra, thiếu SoDonGop
                    worksheet.Cell(b3, COL_NGUOIDI).Value = "tiền lấy";
                    if (detail != null && detail.SoDonGop > 0)
                    {
                        // Có đơn gộp → ghi giá trị đã tính từ UI
                        worksheet.Cell(b3, COL_NGUOILAY).Value = (double)detail.TienLay;
                        decimal donLayThucTe = detail.DonLayThucTe;
                        worksheet.Cell(b3, COL_NGAYLAY).Value = $"{donLayThucTe:N0} đơn";
                    }
                    else
                    {
                        // Không có đơn gộp → formula chính xác = -(SoDon - SoDonTra) × 2k
                        worksheet.Cell(b3, COL_NGUOILAY).FormulaA1 =
                            $"-(COUNTIFS({rShop},\"<>\",{rNguoiDi},\"{nd}\")-COUNTIFS({rShop},\"<>\",{rNguoiDi},\"{nd}\",{rFail},\"xx\"))*{AppConstants.PHI_LAY_HANG_MOI_DON}";
                        if (detail != null)
                        {
                            decimal donLayThucTe = detail.DonLayThucTe;
                            worksheet.Cell(b3, COL_NGAYLAY).Value = $"{donLayThucTe:N0} đơn";
                        }
                    }

                    // ── đơn trả ───────────────────────────────────────────────
                    worksheet.Cell(b4, COL_NGUOIDI).Value = "đơn trả";
                    worksheet.Cell(b4, COL_NGUOIDI).Style.Font.FontColor = XLColor.Red;
                    if (detail != null && detail.SoDonTra > 0)
                    {
                        worksheet.Cell(b4, COL_NGUOILAY).Value = (double)detail.TienDonTra;
                        worksheet.Cell(b4, COL_NGAYLAY).Value = $"{detail.SoDonTra} đơn";
                        worksheet.Cell(b4, COL_NGUOILAY).Style.Font.FontColor = XLColor.Red;
                    }
                    else
                    {
                        worksheet.Cell(b4, COL_NGUOILAY).Value = 0;
                        worksheet.Cell(b4, COL_NGAYLAY).Value = 0;
                    }
                }
                else
                {
                    // An Tâm — không trừ ship/lấy/trả
                    worksheet.Cell(b2, COL_NGUOIDI).Value = "tiền ship";
                    worksheet.Cell(b2, COL_NGUOILAY).Value = "—";
                    worksheet.Cell(b3, COL_NGUOIDI).Value = "tiền lấy";
                    worksheet.Cell(b3, COL_NGUOILAY).Value = "—";
                    worksheet.Cell(b4, COL_NGUOIDI).Value = "đơn trả";
                    worksheet.Cell(b4, COL_NGUOIDI).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(b4, COL_NGUOILAY).Value = "—";
                }

                worksheet.Cell(b5, COL_NGUOIDI).Value = "đơn cũ ck";
                worksheet.Cell(b5, COL_NGUOIDI).Style.Font.FontColor = XLColor.Red;

                // Tổng cuối block — luôn dùng SUM để gộp cả formula lẫn value
                var cTotal = worksheet.Cell(b6, COL_NGUOILAY);
                cTotal.FormulaA1 = $"SUM({nguoiLayColL}{b1}:{nguoiLayColL}{b5})";
                cTotal.Style.Font.Bold = true;
                cTotal.Style.Fill.BackgroundColor = XLColor.LightBlue;

                var cTotalDon = worksheet.Cell(b6, COL_NGAYLAY);
                cTotalDon.FormulaA1 = $"{ngayLayColL}{b1}";
                cTotalDon.Style.Font.Bold = true;
                cTotalDon.Style.Fill.BackgroundColor = XLColor.LightBlue;

                curRow += SUMMARY_RIGHT_BLOCK_HEIGHT + 1;
            }
        }

        /// <summary>
        /// Tạo bảng đối soát gửi shop (per SHOP × NGÀY).
        /// Layout matching template Excel gốc:
        ///   [T7.28.1]       | Tiền     | Số đơn
        ///   Tiền hàng        | SUBTOTAL | count
        ///   Trừ Ship         | 0 (red)  |
        ///   Cước xử          | 0 (red)  |
        ///   Khách C          | 0 (red)  |
        ///   Giảm tiền        | 0        |
        ///   Hàng Boom Trả    |          |
        ///   Trừ đơn cũ đã ck |          |
        ///   Tổng             | formula  | count
        /// </summary>
        private void BuildLeftSummary(
            IXLWorksheet worksheet,
            List<(string Shop, string Ngay)> distinctShopDates,
            int startRow,
            string rThu,
            string rShip,
            string rShop,
            string rNgay,
            string rMa,
            string tenkhColL,
            string diachiColL,
            TongHopSummary tongHop = null
        )
        {
            // Tìm TIỀN HÀNG range cho SUMIFS
            string hangColL = ColLetter(COL_TIENHANG);
            int lastDataRow = DATA_START_ROW;
            foreach (var row in worksheet.RowsUsed())
            {
                int rn = row.RowNumber();
                if (rn < DATA_START_ROW)
                    continue;
                string shopVal = row.Cell(COL_SHOP).GetString();
                if (!string.IsNullOrWhiteSpace(shopVal))
                    lastDataRow = Math.Max(lastDataRow, rn);
            }
            string rHang = $"{hangColL}${DATA_START_ROW}:{hangColL}${lastDataRow}";
            // FAIL range for boom detection
            string failColL = ColLetter(COL_FAIL);
            string rFail = $"{failColL}${DATA_START_ROW}:{failColL}${lastDataRow}";

            int curRow = startRow;

            foreach (var (shop, ngay) in distinctShopDates)
            {
                // Parse ngày để tạo header T7.28.1 format
                // ngay format: "28-02-2026." → "T7.28.2" (Thứ.Ngày.Tháng)
                string headerLabel = FormatDoiSoatHeader(ngay);

                int r0 = curRow; // Header: T7.28.1 | Tiền | Số đơn
                int r1 = curRow + 1; // Tiền hàng | SUBTOTAL | count
                int r2 = curRow + 2; // Trừ Ship (placeholder 0, red)
                int r3 = curRow + 3; // Cước xử (placeholder 0, red)
                int r4 = curRow + 4; // Khách C (placeholder 0, red)
                int r5 = curRow + 5; // Giảm tiền (placeholder 0)
                int r6 = curRow + 6; // Hàng Boom Trả (formula from FAIL=xx)
                int r7 = curRow + 7; // Trừ đơn cũ đã ck (manual)
                int r8 = curRow + 8; // (blank separator)
                int r9 = curRow + 9; // Tổng | SUBTOTAL | count

                string aShop = $"$A${r0}";
                string aNgay = $"$A${r1}";

                // ── R0: Header ────────────────────────────────────────────────
                var cHeader = worksheet.Cell(r0, COL_TINHTRANG);
                cHeader.Value = headerLabel;
                cHeader.Style.Font.Bold = true;
                cHeader.Style.Fill.BackgroundColor = XLColor.LightYellow;
                SetBold(worksheet.Cell(r0, COL_TENKH), "Tiền");
                SetBold(worksheet.Cell(r0, COL_DIACHI), "Số đơn");

                // ── R1: Tiền hàng (= TIỀN THU, giống UI TỔNG HỢP) ──────────
                // Store shop + ngày values for SUMIFS reference
                worksheet.Cell(r0, COL_SHOP).Value = shop; // hidden reference cho SUMIFS
                worksheet.Cell(r1, COL_TINHTRANG).Value = ngay; // hidden reference
                worksheet.Cell(r1, COL_SHOP).Value = "Tiền hàng";
                worksheet.Cell(r1, COL_SHOP).Style.Fill.BackgroundColor = XLColor.LightYellow;
                // UI: "Tiền hàng" = SUM(TIỀN THU) → dùng SUMIFS trên TIỀN THU (không phải TIỀN HÀNG)
                worksheet.Cell(r1, COL_TENKH).FormulaA1 =
                    $"SUMIFS({rThu},{rShop},{ColLetter(COL_SHOP)}${r0},{rNgay},{ColLetter(COL_TINHTRANG)}${r1})";
                worksheet.Cell(r1, COL_TENKH).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(r1, COL_DIACHI).FormulaA1 =
                    $"COUNTIFS({rShop},{ColLetter(COL_SHOP)}${r0},{rNgay},{ColLetter(COL_TINHTRANG)}${r1})";
                worksheet.Cell(r1, COL_DIACHI).Style.Fill.BackgroundColor = XLColor.LightYellow;

                // ── R2: Trừ Ship (= -SUMIFS(TIỀN SHIP), giống UI) ────────────
                worksheet.Cell(r2, COL_SHOP).Value = "Trừ Ship";
                // UI: "Trừ Ship" = -TongTienShip → dùng -SUMIFS trên TIỀN SHIP
                worksheet.Cell(r2, COL_TENKH).FormulaA1 =
                    $"-SUMIFS({rShip},{rShop},{ColLetter(COL_SHOP)}${r0},{rNgay},{ColLetter(COL_TINHTRANG)}${r1})";
                worksheet.Cell(r2, COL_TENKH).Style.Font.FontColor = XLColor.Red;
                worksheet.Cell(r2, COL_SHOP).Style.Font.FontColor = XLColor.Red;
                worksheet.Cell(r2, COL_DIACHI).FormulaA1 =
                    $"COUNTIFS({rShop},{ColLetter(COL_SHOP)}${r0},{rNgay},{ColLetter(COL_TINHTRANG)}${r1})";

                // ── R3: Cước xử (placeholder 0, tô đỏ) ──────────────────────
                worksheet.Cell(r3, COL_SHOP).Value = "Cước xử";
                worksheet.Cell(r3, COL_TENKH).Value = 0;
                worksheet.Cell(r3, COL_TENKH).Style.Font.FontColor = XLColor.Red;
                worksheet.Cell(r3, COL_SHOP).Style.Font.FontColor = XLColor.Red;

                // ── R4: Khách C (placeholder 0, tô đỏ) ──────────────────────
                worksheet.Cell(r4, COL_SHOP).Value = "Khách Chuyển Khoản";
                worksheet.Cell(r4, COL_TENKH).Value = 0;
                worksheet.Cell(r4, COL_TENKH).Style.Font.FontColor = XLColor.Red;
                worksheet.Cell(r4, COL_SHOP).Style.Font.FontColor = XLColor.Red;

                // ── R5: Giảm tiền (placeholder 0) ────────────────────────────
                worksheet.Cell(r5, COL_SHOP).Value = "Giảm tiền";
                worksheet.Cell(r5, COL_TENKH).Value = 0;

                // ── R6: Hàng Boom Trả (auto-filled từ UI data) ──────────────
                worksheet.Cell(r6, COL_SHOP).Value = "Hàng Boom Trả";
                if (tongHop != null && tongHop.SoDonTra > 0)
                {
                    worksheet.Cell(r6, COL_TENKH).Value = (double)tongHop.TienDonTra;
                    worksheet.Cell(r6, COL_DIACHI).Value = tongHop.SoDonTra;
                    worksheet.Cell(r6, COL_TENKH).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(r6, COL_SHOP).Style.Font.FontColor = XLColor.Red;
                }

                // ── R7: Trừ đơn cũ đã ck (auto-filled từ UI data) ──────────
                worksheet.Cell(r7, COL_SHOP).Value =
                    tongHop != null
                    && tongHop.SoRowNegative == 1
                    && !string.IsNullOrEmpty(tongHop.NegativeLabel)
                        ? tongHop.NegativeLabel
                        : "Trừ đơn cũ đã ck";
                worksheet.Cell(r7, COL_SHOP).Style.Fill.BackgroundColor = XLColor.FromArgb(
                    255,
                    200,
                    124
                );
                if (tongHop != null && tongHop.SoRowNegative > 0)
                {
                    worksheet.Cell(r7, COL_TENKH).Value = (double)tongHop.TongNegative;
                    worksheet.Cell(r7, COL_DIACHI).Value = $"{tongHop.SoRowNegative} dòng";
                    worksheet.Cell(r7, COL_TENKH).Style.Font.FontColor = XLColor.FromArgb(
                        200,
                        100,
                        0
                    );
                }

                // ── R9: Tổng (formula = sum r1..r7, giống UI) ────────────
                var cTong = worksheet.Cell(r9, COL_SHOP);
                cTong.Value = "Tổng";
                cTong.Style.Font.Bold = true;
                cTong.Style.Fill.BackgroundColor = XLColor.LightYellow;
                var cTongVal = worksheet.Cell(r9, COL_TENKH);
                cTongVal.FormulaA1 = $"SUM({tenkhColL}{r1}:{tenkhColL}{r7})";
                cTongVal.Style.Font.Bold = true;
                cTongVal.Style.Fill.BackgroundColor = XLColor.LightYellow;
                var cTongDon = worksheet.Cell(r9, COL_DIACHI);
                cTongDon.FormulaA1 = $"{diachiColL}{r1}"; // same as Tiền hàng Số đơn
                cTongDon.Style.Font.Bold = true;
                cTongDon.Style.Fill.BackgroundColor = XLColor.LightYellow;

                // Nền LightCyan cho toàn block
                for (int r = r0; r <= r9; r++)
                for (int c = COL_TINHTRANG; c <= COL_QUAN; c++)
                {
                    var bg = worksheet.Cell(r, c).Style.Fill.BackgroundColor;
                    if (bg.Equals(XLColor.NoColor) || bg.Equals(XLColor.White))
                        worksheet.Cell(r, c).Style.Fill.BackgroundColor = XLColor.LightCyan;
                }

                // Restore màu ưu tiên
                cHeader.Style.Fill.BackgroundColor = XLColor.LightYellow;
                cTong.Style.Fill.BackgroundColor = XLColor.LightYellow;
                cTongVal.Style.Fill.BackgroundColor = XLColor.LightYellow;

                curRow += SUMMARY_LEFT_BLOCK_HEIGHT + 1;
            }
        }

        /// <summary>
        /// Format đối soát header: "28-02-2026." → "T7.28.2" (Thứ.Ngày.Tháng)
        /// </summary>
        private static string FormatDoiSoatHeader(string ngayStr)
        {
            string cleaned = ngayStr.TrimEnd('.');
            if (
                DateTime.TryParseExact(
                    cleaned,
                    new[] { "dd-MM-yyyy", "d-M-yyyy", "dd/MM/yyyy", "d/M/yyyy", "dd-MM", "d-M" },
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    out DateTime dt
                )
            )
            {
                string thu = dt.DayOfWeek switch
                {
                    DayOfWeek.Monday => "T2",
                    DayOfWeek.Tuesday => "T3",
                    DayOfWeek.Wednesday => "T4",
                    DayOfWeek.Thursday => "T5",
                    DayOfWeek.Friday => "T6",
                    DayOfWeek.Saturday => "T7",
                    DayOfWeek.Sunday => "CN",
                    _ => "T?",
                };
                return $"{thu}.{dt.Day}.{dt.Month}";
            }
            // Fallback: giữ nguyên
            return ngayStr;
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

        // ── Cập nhật nhiều đơn hàng theo danh sách MÃ ──────────────────────────

        /// <summary>
        /// Danh sách tên cột có thể chỉnh sửa trong tab "Cập nhật".
        /// </summary>
        public static readonly string[] EditableColumnNames = new[]
        {
            "TÌNH TRẠNG TT",
            "SHOP",
            "TÊN KH",
            "ĐỊA CHỈ",
            "QUẬN",
            "TIỀN THU",
            "TIỀN SHIP",
            "NGƯỜI ĐI",
            "NGƯỜI LẤY",
            "NGÀY LẤY",
            "GHI CHÚ",
            "ỨNG TIỀN",
            "HÀNG TỒN",
            "FAIL",
        };

        /// <summary>
        /// Cập nhật các field cho nhiều đơn hàng theo danh sách MÃ HĐ.
        /// Trả về (số đơn đã cập nhật, danh sách mã không tìm thấy).
        /// </summary>
        public (int updated, List<string> notFound) UpdateInvoiceFields(
            string sheetName,
            List<string> maList,
            Dictionary<string, string> fieldValues
        )
        {
            int updated = 0;
            var notFound = new List<string>();

            using var workbook = new XLWorkbook(_excelFilePath);
            var worksheet = workbook.Worksheet(sheetName);

            // Build column index map từ header row
            var usedRange = worksheet.RangeUsed();
            if (usedRange == null)
                return (0, maList.ToList());

            int rowCount = usedRange.RowCount();
            int colCount = usedRange.ColumnCount();

            // Tìm header row
            int headerRow = -1;
            for (int r = 1; r <= Math.Min(5, rowCount); r++)
            {
                for (int c = 1; c <= Math.Min(colCount, 5); c++)
                {
                    string cv = worksheet.Cell(r, c).GetString()?.Trim() ?? "";
                    if (
                        cv == "SHOP"
                        || cv.Contains("Tình trạng", StringComparison.OrdinalIgnoreCase)
                        || cv.Contains("TIỀN THU", StringComparison.OrdinalIgnoreCase)
                        || cv.Contains("TÊN KH", StringComparison.OrdinalIgnoreCase)
                    )
                    {
                        headerRow = r;
                        break;
                    }
                }
                if (headerRow > 0)
                    break;
            }
            if (headerRow < 0)
                return (0, maList.ToList());

            // Map tên cột → index
            var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = 1; c <= colCount; c++)
            {
                string h = worksheet.Cell(headerRow, c).GetString()?.Trim() ?? "";
                if (!string.IsNullOrEmpty(h) && !colMap.ContainsKey(h))
                    colMap[h] = c;
            }

            // Ghi các field
            foreach (string ma in maList)
            {
                int targetRow = FindRowByMa(worksheet, ma);
                if (targetRow < 0)
                {
                    notFound.Add(ma);
                    continue;
                }
                foreach (var kv in fieldValues)
                {
                    if (!colMap.TryGetValue(kv.Key, out int colIdx))
                        continue;
                    var cell = worksheet.Cell(targetRow, colIdx);
                    if (cell.HasFormula)
                        continue; // không ghi đè formula
                    if (decimal.TryParse(kv.Value, out decimal num))
                        cell.SetValue(num);
                    else
                        cell.SetValue(kv.Value);
                }
                updated++;
            }

            workbook.Save();
            return (updated, notFound);
        }
    }
}

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
        private const int SUMMARY_LEFT_BLOCK_HEIGHT = 9;

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

                    // Upsert: tìm row có MÃ trùng → ghi đè; nếu MÃ rỗng → thử match TÊN KH (hàng sỉ ghi đè theo tên)
                    int targetRow;
                    if (!isMissing)
                    {
                        targetRow = FindRowByMa(worksheet, ma);
                    }
                    else
                    {
                        // Hàng sỉ (không có MÃ HĐ) → ghi đè dựa trên TÊN KH
                        string tenKH = data.GetValueOrDefault("TÊN KH", "");
                        targetRow = string.IsNullOrWhiteSpace(tenKH)
                            ? -1
                            : FindRowByTenKH(worksheet, tenKH);
                    }
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
                // AT hôm nay – dùng để phân biệt AT ngày cũ (đơn trả) vs AT hôm nay
                string atTodayForDistinct =
                    AppConstants.NGUOI_DI_DEFAULT + DateTime.Now.ToString("dd-MM");
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
                    {
                        // Không phải shipper thật (vd: "lưu trả") → skip khỏi per-shipper summary
                        bool isNotShipper = AppConstants.NOT_SHIPPER_VALUES.Any(v =>
                            nguoiDi.Contains(v, StringComparison.OrdinalIgnoreCase)
                        );
                        // AT ngày cũ (vd "AT 30-03" khi hôm nay là 08-04) → đơn trả, skip
                        bool isATOldDate =
                            nguoiDi.StartsWith(
                                AppConstants.NGUOI_DI_DEFAULT,
                                StringComparison.OrdinalIgnoreCase
                            )
                            && !nguoiDi.StartsWith(
                                atTodayForDistinct,
                                StringComparison.OrdinalIgnoreCase
                            );
                        if (!isNotShipper && !isATOldDate)
                            distinctNguoiDis.Add(nguoiDi);
                    }
                    if (!string.IsNullOrWhiteSpace(shop) && !string.IsNullOrWhiteSpace(ngay))
                    {
                        // Skip hàng tồn (carry-over ngày trước) — không tạo LEFT block riêng
                        string hangTonVal = worksheet
                            .Cell(r, COL_HANGTON)
                            .GetString()
                            .Trim()
                            .ToLower();
                        if (hangTonVal != "x")
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
                string hangColL = ColLetter(COL_TIENHANG);
                string rHang = $"{hangColL}${DATA_START_ROW}:{hangColL}${lastDataRow}";
                string failColL = ColLetter(COL_FAIL);
                string rFail = $"{failColL}${DATA_START_ROW}:{failColL}${lastDataRow}";
                string ungColL = ColLetter(COL_UNGIEN);
                string rUng = $"{ungColL}${DATA_START_ROW}:{ungColL}${lastDataRow}";
                string col1ColL = ColLetter(COL_COL1);
                string rCol1 = $"{col1ColL}${DATA_START_ROW}:{col1ColL}${lastDataRow}";
                string ghichuColL = ColLetter(COL_GHICHU);
                string rGhiChu = $"{ghichuColL}${DATA_START_ROW}:{ghichuColL}${lastDataRow}";

                // ── Auto-detect đơn gộp → ghi GHI CHÚ = "gộp" ──────────────
                AutoMarkDonGop(worksheet, lastDataRow);

                // ── AT zone breakdown rows (cột H, giữa subtotal và summary) ──
                // Tính số đơn AT theo zone phí ship, ghi =-count*fee vào cột H
                string atNguoiDi = distinctNguoiDis.FirstOrDefault(n =>
                    n.StartsWith(AppConstants.NGUOI_DI_DEFAULT, StringComparison.OrdinalIgnoreCase)
                );
                int atZoneStartRow = -1,
                    atZoneEndRow = -1;
                if (!string.IsNullOrEmpty(atNguoiDi))
                {
                    var atZoneCounts = new Dictionary<decimal, int>();
                    for (int r = DATA_START_ROW; r <= lastDataRow; r++)
                    {
                        string shopVal = worksheet.Cell(r, COL_SHOP).GetString().Trim();
                        if (string.IsNullOrWhiteSpace(shopVal))
                            continue;
                        string nguoi = worksheet.Cell(r, COL_NGUOIDI).GetString().Trim();
                        // Chỉ match AT hôm nay (exact), không match AT ngày cũ
                        if (!nguoi.Equals(atNguoiDi, StringComparison.OrdinalIgnoreCase))
                            continue;
                        string quan = worksheet.Cell(r, COL_QUAN).GetString().Trim();
                        decimal atFee = LookupShipFeeByDict(quan, AppConstants.AT_SHIPPING_FEES);
                        if (atFee == 0m)
                            continue;
                        if (!atZoneCounts.ContainsKey(atFee))
                            atZoneCounts[atFee] = 0;
                        atZoneCounts[atFee]++;
                    }
                    if (atZoneCounts.Count > 0)
                    {
                        int zoneRow = subtotalRow + 1;
                        atZoneStartRow = zoneRow;
                        foreach (var zone in atZoneCounts.OrderBy(z => z.Key))
                        {
                            worksheet.Cell(zoneRow, COL_TIENSHIP).FormulaA1 =
                                $"-{zone.Value}*{zone.Key:0}";
                            zoneRow++;
                        }
                        atZoneEndRow = zoneRow - 1;
                    }
                }

                // ── BẢNG PHẢI: per NGƯỜI ĐI ───────────────────────────────────
                int rightEndRow = BuildRightSummary(
                    worksheet,
                    distinctNguoiDis,
                    summaryRow,
                    subtotalRow,
                    lastDataRow,
                    atZoneStartRow,
                    atZoneEndRow,
                    rThu,
                    rShip,
                    rNguoiDi,
                    rMa,
                    rShop,
                    rCol1,
                    rGhiChu
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
                    rHang,
                    rFail,
                    rUng,
                    rCol1
                );

                // ── BẢNG LỢI NHUẬN ──────────────────────────────────────────
                BuildProfitSummary(
                    worksheet,
                    distinctNguoiDis,
                    summaryRow,
                    rThu,
                    rNguoiDi,
                    rShop,
                    tenkhColL
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
        /// Tìm row có TÊN KH trùng (case-insensitive) trong sheet.
        /// Dùng để ghi đè đơn hàng sỉ (không có MÃ HĐ) — match dựa trên tên khách hàng.
        /// Trả về -1 nếu không tìm thấy.
        /// </summary>
        private int FindRowByTenKH(IXLWorksheet worksheet, string tenKH)
        {
            if (string.IsNullOrWhiteSpace(tenKH))
                return -1;
            string target = tenKH.Trim();
            foreach (var row in worksheet.RowsUsed())
            {
                if (row.RowNumber() <= 2)
                    continue;
                string existing = row.Cell(COL_TENKH).GetString().Trim();
                if (existing.Equals(target, StringComparison.OrdinalIgnoreCase))
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

            // TIỀN HÀNG = TIỀN THU - TIỀN SHIP (luôn dùng 1 công thức duy nhất)
            string thuCol = ColLetter(COL_TIENTHU);
            string shipCol = ColLetter(COL_TIENSHIP);
            worksheet.Cell(targetRow, COL_TIENHANG).FormulaA1 =
                $"{thuCol}{targetRow}-{shipCol}{targetRow}";

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

        /// <summary>
        /// Auto-detect đơn gộp (cùng TÊN KH + ĐỊA CHỈ) và ghi GHI CHÚ = "gộp".
        /// Đồng thời ghi cột COL1 (Q) = 1 cho mỗi data row (dùng cho COUNTIFS).
        /// </summary>
        private void AutoMarkDonGop(IXLWorksheet worksheet, int lastDataRow)
        {
            // Collect (tenKH, diaChi) → list of rows
            var groups = new Dictionary<string, List<int>>();
            for (int r = DATA_START_ROW; r <= lastDataRow; r++)
            {
                string shop = worksheet.Cell(r, COL_SHOP).GetString().Trim();
                if (string.IsNullOrWhiteSpace(shop))
                    continue;
                string tenKH = worksheet.Cell(r, COL_TENKH).GetString().Trim();
                string diaChi = worksheet.Cell(r, COL_DIACHI).GetString().Trim();
                if (string.IsNullOrEmpty(tenKH) || string.IsNullOrEmpty(diaChi))
                    continue;
                string key = $"{tenKH.ToLower()}|{diaChi.ToLower()}";
                if (!groups.ContainsKey(key))
                    groups[key] = new List<int>();
                groups[key].Add(r);
            }

            // Mark gộp groups
            foreach (var kv in groups)
            {
                if (kv.Value.Count <= 1)
                    continue;
                foreach (int r in kv.Value)
                {
                    string existing = worksheet.Cell(r, COL_GHICHU).GetString().Trim();
                    if (string.IsNullOrEmpty(existing))
                        worksheet.Cell(r, COL_GHICHU).Value = "gộp";
                    else if (!existing.Contains("gộp", StringComparison.OrdinalIgnoreCase))
                        worksheet.Cell(r, COL_GHICHU).Value = $"gộp, {existing}";
                }
            }

            // Write COL1 (Q) = 1 for each data row (used for COUNTIFS in summary)
            for (int r = DATA_START_ROW; r <= lastDataRow; r++)
            {
                string shop = worksheet.Cell(r, COL_SHOP).GetString().Trim();
                if (!string.IsNullOrWhiteSpace(shop))
                    worksheet.Cell(r, COL_COL1).Value = 1;
            }
        }

        /// <summary>Tạo bảng tổng kết bên phải (per NGƯỜI ĐI), mỗi người 1 block 7 dòng.</summary>
        private int BuildRightSummary(
            IXLWorksheet worksheet,
            List<string> distinctNguoiDis,
            int startRow,
            int subtotalRow,
            int lastDataRow,
            int atZoneStartRow,
            int atZoneEndRow,
            string rThu,
            string rShip,
            string rNguoiDi,
            string rMa,
            string rShop,
            string rCol1,
            string rGhiChu
        )
        {
            // RIGHT summary columns matching template: K(11)=labels, L(12)=values, M(13)=counts
            string valColL = ColLetter(COL_NGAYLAY); // L — values (Tiền Thu)
            string cntColL = ColLetter(COL_GHICHU); // M — counts (Số đơn)
            string nameColL = ColLetter(COL_NGUOILAY); // K — labels/names
            string shipHColL = ColLetter(COL_TIENSHIP); // H — ship fee / zone calcs
            string col1ColL = ColLetter(COL_COL1); // Q — order count column

            // ── Thu thập data rows per NGƯỜI ĐI để tính đơn gộp + đơn trả ──
            int computedLastDataRow = DATA_START_ROW - 1;
            foreach (var row in worksheet.RowsUsed())
            {
                int rn = row.RowNumber();
                if (rn < DATA_START_ROW)
                    continue;
                string shopVal = row.Cell(COL_SHOP).GetString().Trim();
                if (!string.IsNullOrWhiteSpace(shopVal))
                    computedLastDataRow = Math.Max(computedLastDataRow, rn);
            }

            // Collect per-person row data
            var rowsPerNguoi = new Dictionary<
                string,
                List<(
                    string TenKH,
                    string DiaChi,
                    double TienThu,
                    double ShipFee,
                    string Quan,
                    bool IsTra
                )>
            >(StringComparer.OrdinalIgnoreCase);

            // Pre-collected đơn trả AT ngày cũ → tính vào AT hôm nay
            string atToday = AppConstants.NGUOI_DI_DEFAULT + DateTime.Now.ToString("dd-MM");
            double atOldDateDonTra = 0;
            int atOldDateDonTraCount = 0;

            for (int r = DATA_START_ROW; r <= computedLastDataRow; r++)
            {
                string shopVal = worksheet.Cell(r, COL_SHOP).GetString().Trim();
                if (string.IsNullOrWhiteSpace(shopVal))
                    continue;
                string nguoi = worksheet.Cell(r, COL_NGUOIDI).GetString().Trim();
                if (string.IsNullOrEmpty(nguoi))
                    nguoi = "(không rõ)";

                string tenKH = worksheet.Cell(r, COL_TENKH).GetString().Trim();
                string diaChi = worksheet.Cell(r, COL_DIACHI).GetString().Trim();
                string quan = worksheet.Cell(r, COL_QUAN).GetString().Trim();
                double.TryParse(worksheet.Cell(r, COL_TIENTHU).GetString(), out double tienThu);
                double.TryParse(worksheet.Cell(r, COL_TIENSHIP).GetString(), out double shipFee);
                string failVal = worksheet.Cell(r, COL_FAIL).GetString().Trim().ToLower();
                bool isTra = failVal.Contains("xx");

                // AT ngày cũ (VD: "AT 30-03" khi hôm nay 08-04) → đơn trả
                // Tính tiền trừ vào AT hôm nay, sửa NGƯỜI ĐI thành "lưu trả"
                bool isATRow = nguoi.StartsWith(
                    AppConstants.NGUOI_DI_DEFAULT,
                    StringComparison.OrdinalIgnoreCase
                );
                bool isATOldDate =
                    isATRow && !nguoi.StartsWith(atToday, StringComparison.OrdinalIgnoreCase);

                if (isATOldDate)
                {
                    // Tính đơn trả: -(tiền thu), KHÔNG có phí công 5k
                    atOldDateDonTra += (double)(-(decimal)tienThu);
                    atOldDateDonTraCount++;

                    // Sửa NGƯỜI ĐI thành "lưu trả" trong Excel
                    worksheet.Cell(r, COL_NGUOIDI).Value = "lưu trả";
                    continue; // Skip — không tính vào report nhỏ
                }

                if (!rowsPerNguoi.ContainsKey(nguoi))
                    rowsPerNguoi[nguoi] =
                        new List<(string, string, double, double, string, bool)>();
                rowsPerNguoi[nguoi].Add((tenKH, diaChi, tienThu, shipFee, quan, isTra));
            }

            int curRow = startRow;

            foreach (string nd in distinctNguoiDis)
            {
                // Tính đơn gộp + đơn trả từ data
                int soDon = 0,
                    soDonGop = 0,
                    soDonTra = 0;
                double totalTienDonTra = 0;
                if (rowsPerNguoi.TryGetValue(nd, out var rows))
                {
                    soDon = rows.Count;
                    soDonTra = rows.Count(r => r.IsTra);

                    var groups = rows.Where(r =>
                            !string.IsNullOrEmpty(r.TenKH) && !string.IsNullOrEmpty(r.DiaChi)
                        )
                        .GroupBy(r => (r.TenKH.ToLower(), r.DiaChi.ToLower()));
                    foreach (var g in groups)
                        if (g.Count() > 1)
                            soDonGop += g.Count() - 1;

                    foreach (var r in rows.Where(r => r.IsTra))
                    {
                        // AT dùng AT_SHIPPING_FEES (zone riêng), shipper khác dùng SHIPPING_FEES_BY_QUAN
                        bool isAT = nd.StartsWith(
                            AppConstants.NGUOI_DI_DEFAULT,
                            StringComparison.OrdinalIgnoreCase
                        );
                        decimal shipFeeLookup = isAT
                            ? LookupShipFeeByDict(r.Quan, AppConstants.AT_SHIPPING_FEES)
                            : LookupShipFeeByQuan(r.Quan);
                        // AT: trừ đúng tiền thu (hàng + ship AT), KHÔNG có phí công 5k
                        // Khác: trừ tiền hàng + phí công 5k
                        totalTienDonTra += isAT
                            ? (double)(-(decimal)r.TienThu)
                            : (double)(
                                -(decimal)r.TienThu + shipFeeLookup - AppConstants.PHI_CONG_DON_TRA
                            );
                    }
                }

                // Cộng thêm đơn trả AT ngày cũ (đã sửa thành "lưu trả" ở trên)
                bool isAnTam = nd.StartsWith(
                    AppConstants.NGUOI_DI_DEFAULT,
                    StringComparison.OrdinalIgnoreCase
                );
                if (isAnTam && atOldDateDonTraCount > 0)
                {
                    totalTienDonTra += atOldDateDonTra;
                    soDonTra += atOldDateDonTraCount;
                }
                int soDonGiao = soDon - soDonGop;

                bool isNguoiLay = nd.Equals(
                    AppConstants.NGUOI_LAY_DEFAULT,
                    StringComparison.OrdinalIgnoreCase
                );

                int b0 = curRow;
                int b1 = curRow + 1;
                int b2 = curRow + 2;
                int b3 = curRow + 3;
                int b4 = curRow + 4;
                int b5 = curRow + 5;
                int b6 = curRow + 6;

                // Cell reference for SUMIFS: name cell in K column
                string nameRef = $"{nameColL}${b0}";

                // Header block — cols K,L,M
                SetBoldYellow(worksheet.Cell(b0, COL_NGUOILAY), nd, XLColor.LightSteelBlue);
                SetBoldYellow(worksheet.Cell(b0, COL_NGAYLAY), "Tiền Thu", XLColor.LightSteelBlue);
                SetBoldYellow(worksheet.Cell(b0, COL_GHICHU), "Số đơn", XLColor.LightSteelBlue);

                // TỔNG ĐƠN NHẬN — SUMIFS with cell reference (matching template)
                worksheet.Cell(b1, COL_NGUOILAY).Value = "TỔNG ĐƠN NHẬN";
                worksheet.Cell(b1, COL_NGAYLAY).FormulaA1 = $"SUMIFS({rThu},{rNguoiDi},{nameRef})";
                worksheet.Cell(b1, COL_GHICHU).FormulaA1 = $"SUMIFS({rCol1},{rNguoiDi},{nameRef})";

                // tiền ship
                worksheet.Cell(b2, COL_NGUOILAY).Value = "tiền ship";
                if (isAnTam && atZoneStartRow > 0 && atZoneEndRow > 0)
                {
                    // AT: tiền ship = SUM of zone breakdown rows in column H
                    worksheet.Cell(b2, COL_NGAYLAY).FormulaA1 =
                        $"SUM({shipHColL}{atZoneStartRow}:{shipHColL}{atZoneEndRow})";
                    // Số đơn cho AT: tổng đơn (không trừ gộp, vì zone tính per-order)
                    worksheet.Cell(b2, COL_GHICHU).FormulaA1 =
                        $"SUMIFS({rCol1},{rNguoiDi},{nameRef})";
                }
                else
                {
                    // Shipper khác: -(SUMIFS(Ship) - SoDonGiao × 5k)
                    worksheet.Cell(b2, COL_GHICHU).FormulaA1 =
                        $"SUMIFS({rCol1},{rNguoiDi},{nameRef})-COUNTIFS({rNguoiDi},{nameRef},{rGhiChu},\"*gộp*\")";
                    worksheet.Cell(b2, COL_NGAYLAY).FormulaA1 =
                        $"-SUMIFS({rShip},{rNguoiDi},{nameRef})+{cntColL}{b2}*{AppConstants.PHI_SHIP_MOI_DON}";
                }

                // tiền lấy — chỉ có giá trị cho NGUOI_LAY_DEFAULT (c.cuong)
                worksheet.Cell(b3, COL_NGUOILAY).Value = "tiền lấy";
                if (isNguoiLay)
                {
                    // M = totalOrders - COUNTIFS("*gộp*")/2 - COUNTIFS(FAIL,"*xx*")
                    // Đơn trả (FAIL=xx) không tính tiền lấy vì không "lấy" thực sự.
                    string rGhiChuFull =
                        $"{ColLetter(COL_GHICHU)}${DATA_START_ROW}:{ColLetter(COL_GHICHU)}${lastDataRow}";
                    string rFailFull =
                        $"{ColLetter(COL_FAIL)}${DATA_START_ROW}:{ColLetter(COL_FAIL)}${lastDataRow}";
                    worksheet.Cell(b3, COL_GHICHU).FormulaA1 =
                        $"{col1ColL}{subtotalRow}-(COUNTIFS({rGhiChuFull},\"*gộp*\")/2)-COUNTIFS({rFailFull},\"*xx*\")*2";
                    worksheet.Cell(b3, COL_NGAYLAY).FormulaA1 =
                        $"-{cntColL}{b3}*{AppConstants.PHI_LAY_HANG_MOI_DON}";
                }

                // đơn trả — auto-filled from FAIL=xx data
                worksheet.Cell(b4, COL_NGUOILAY).Value = "đơn trả";
                worksheet.Cell(b4, COL_NGUOILAY).Style.Font.FontColor = XLColor.Red;
                if (soDonTra > 0)
                {
                    worksheet.Cell(b4, COL_NGAYLAY).Value = totalTienDonTra;
                    worksheet.Cell(b4, COL_NGAYLAY).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(b4, COL_GHICHU).Value = $"{soDonTra} đơn";
                }

                worksheet.Cell(b5, COL_NGUOILAY).Value = "đơn cũ ck";
                worksheet.Cell(b5, COL_NGUOILAY).Style.Font.FontColor = XLColor.Red;

                // Tổng cuối block — SUBTOTAL in L, ref count from L
                var cTotal = worksheet.Cell(b6, COL_NGAYLAY);
                cTotal.FormulaA1 = $"SUBTOTAL(9,{valColL}{b1}:{valColL}{b5})";
                cTotal.Style.Font.Bold = true;
                cTotal.Style.Fill.BackgroundColor = XLColor.LightBlue;

                var cTotalDon = worksheet.Cell(b6, COL_GHICHU);
                cTotalDon.FormulaA1 = $"{cntColL}{b1}";
                cTotalDon.Style.Font.Bold = true;
                cTotalDon.Style.Fill.BackgroundColor = XLColor.LightBlue;

                curRow += SUMMARY_RIGHT_BLOCK_HEIGHT + 1;
            }
            return curRow;
        }

        /// <summary>
        /// Tạo bảng đối soát gửi shop (per SHOP × NGÀY).
        /// Layout matching template Excel thủ công:
        ///   [T7.28.1]          | Shop      | Tiền     | Số đơn
        ///   ngày   | cod       | SUMIFS    | count
        ///          | Trừ Tiền Ship | -SUMIFS | count  (red)
        ///          | Đơn trả &amp; c.khoản |       | (red)
        ///          | Tiền Hàng | SUM       | count
        ///          | đền đơn   |           |
        ///          | nợ cũ     |           |        (cam)
        ///          |           |           |
        ///          | THANH TOÁN| SUM       | CK Đủ 100%
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
            string rHang,
            string rFail,
            string rUng,
            string rCol1
        )
        {
            int curRow = startRow;

            foreach (var (shop, ngay) in distinctShopDates)
            {
                int r0 = curRow; // Header: Shop | Tiền | Số đơn
                int r1 = curRow + 1; // Tiền thu (cod) | SUMIFS(TIỀN THU) | COUNTIFS
                int r2 = curRow + 2; // Trừ Tiền Ship | -SUMIFS(TIỀN SHIP) [red]
                int r3 = curRow + 3; // Đơn trả & c.khoản [red]
                int r4 = curRow + 4; // Tiền Hàng = r1+r2+r3
                int r5 = curRow + 5; // đền đơn (placeholder)
                int r6 = curRow + 6; // nợ cũ (placeholder, cam)
                int r7 = curRow + 7; // (blank separator)
                int r8 = curRow + 8; // THANH TOÁN = SUBTOTAL(r4:r6)

                // ── R0: Header ────────────────────────────────────────────────
                var cHeader = worksheet.Cell(r0, COL_TINHTRANG);
                cHeader.Value = shop; // Shop name in col A (matching template)
                cHeader.Style.Font.Bold = true;
                cHeader.Style.Fill.BackgroundColor = XLColor.LightYellow;
                SetBold(worksheet.Cell(r0, COL_TENKH), "Tiền");
                SetBold(worksheet.Cell(r0, COL_DIACHI), "Số đơn");

                // ── R1: Tiền hàng / cod (SUMIFS TIỀN HÀNG per shop × ngày) ────
                worksheet.Cell(r1, COL_TINHTRANG).Value = ngay; // hidden reference
                worksheet.Cell(r1, COL_SHOP).Value = "cod";
                worksheet.Cell(r1, COL_SHOP).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(r1, COL_TENKH).FormulaA1 =
                    $"SUMIFS({rHang},{rShop},{ColLetter(COL_TINHTRANG)}${r0},{rNgay},{ColLetter(COL_TINHTRANG)}${r1})";
                worksheet.Cell(r1, COL_TENKH).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(r1, COL_DIACHI).FormulaA1 =
                    $"SUMIFS({rCol1},{rShop},{ColLetter(COL_TINHTRANG)}${r0},{rNgay},{ColLetter(COL_TINHTRANG)}${r1})";
                worksheet.Cell(r1, COL_DIACHI).Style.Fill.BackgroundColor = XLColor.LightYellow;

                // ── R2: Trừ Ship (placeholder = 0, ship đã trừ sẵn trong TIỀN HÀNG) ─
                worksheet.Cell(r2, COL_SHOP).Value = "Trừ Ship";
                worksheet.Cell(r2, COL_SHOP).Style.Font.FontColor = XLColor.Red;
                worksheet.Cell(r2, COL_TENKH).Value = 0;
                worksheet.Cell(r2, COL_TENKH).Style.Font.FontColor = XLColor.Red;

                // ── R3: Đơn trả & c.khoản — uses TIỀN HÀNG (not THU), FAIL="xx" + ỨNG TIỀN="x"
                worksheet.Cell(r3, COL_SHOP).Value = "Đơn trả & c.khoản";
                worksheet.Cell(r3, COL_SHOP).Style.Font.FontColor = XLColor.Red;
                worksheet.Cell(r3, COL_TENKH).FormulaA1 =
                    $"-SUMIFS({rHang},{rShop},{ColLetter(COL_TINHTRANG)}${r0},{rFail},\"xx\",{rUng},\"x\")";
                worksheet.Cell(r3, COL_TENKH).Style.Font.FontColor = XLColor.Red;
                worksheet.Cell(r3, COL_DIACHI).FormulaA1 =
                    $"SUMIFS({rCol1},{rShop},{ColLetter(COL_TINHTRANG)}${r0},{rFail},\"xx\",{rUng},\"x\")";
                worksheet.Cell(r3, COL_DIACHI).Style.Font.FontColor = XLColor.Red;

                // ── R4: Tiền Hàng Hcm = SUBTOTAL(9, r1:r3) ─────────────────
                worksheet.Cell(r4, COL_SHOP).Value = "Tiền Hàng Hcm";
                worksheet.Cell(r4, COL_SHOP).Style.Font.Bold = true;
                worksheet.Cell(r4, COL_TENKH).FormulaA1 =
                    $"SUBTOTAL(9,{tenkhColL}{r1}:{tenkhColL}{r3})";
                worksheet.Cell(r4, COL_TENKH).Style.Font.Bold = true;
                worksheet.Cell(r4, COL_DIACHI).FormulaA1 = $"{diachiColL}{r1}";

                // ── R5: đền đơn (placeholder, red) ──────────────────────────
                worksheet.Cell(r5, COL_SHOP).Value = "đền đơn";
                worksheet.Cell(r5, COL_SHOP).Style.Font.FontColor = XLColor.Red;

                // ── R6: nợ cũ (placeholder, tô cam) ─────────────────────────
                worksheet.Cell(r6, COL_SHOP).Value = "nợ cũ";
                worksheet.Cell(r6, COL_SHOP).Style.Fill.BackgroundColor = XLColor.FromArgb(
                    255,
                    200,
                    124
                );

                // ── R8: THANH TOÁN = r4+r5+r6 ─────────────────────────
                var cTong = worksheet.Cell(r8, COL_SHOP);
                cTong.Value = "THANH TOÁN";
                cTong.Style.Font.Bold = true;
                cTong.Style.Fill.BackgroundColor = XLColor.LightYellow;
                var cTongVal = worksheet.Cell(r8, COL_TENKH);
                cTongVal.FormulaA1 = $"{tenkhColL}{r4}+{tenkhColL}{r5}+{tenkhColL}{r6}";
                cTongVal.Style.Font.Bold = true;
                cTongVal.Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(r8, COL_DIACHI).Value = "CK Đủ 100%";
                worksheet.Cell(r8, COL_DIACHI).Style.Font.Bold = true;
                worksheet.Cell(r8, COL_DIACHI).Style.Fill.BackgroundColor = XLColor.LightYellow;

                // Nền LightCyan cho toàn block
                for (int r = r0; r <= r8; r++)
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

                // Merge C:D for value cells (matching template)
                foreach (int mr in new[] { r0, r1, r2, r3, r4, r5, r6, r8 })
                    worksheet.Range(mr, COL_TENKH, mr, COL_MA).Merge();

                // "Sài Gòn" label in F at cod row (matching template)
                worksheet.Cell(r1, COL_QUAN).Value = "Sài Gòn";

                curRow += SUMMARY_LEFT_BLOCK_HEIGHT + 1;
            }
        }

        /// <summary>
        /// Tạo bảng lợi nhuận (cols O-P) — tổng kết per NGƯỜI ĐI + shop + tổng.
        /// Matching reference file layout.
        /// </summary>
        private void BuildProfitSummary(
            IXLWorksheet worksheet,
            List<string> distinctNguoiDis,
            int startRow,
            string rThu,
            string rNguoiDi,
            string rShop,
            string tenkhColL
        )
        {
            // Profit starts at same row as RIGHT summary header, cols O-P (COL_HANGTON, COL_FAIL)
            int headerRow = startRow;
            worksheet.Cell(headerRow, COL_HANGTON).Value = "lợi nhuận";
            worksheet.Cell(headerRow, COL_HANGTON).Style.Font.Bold = true;

            int curRow = headerRow + 1;
            int firstValueRow = curRow;

            // Per NGƯỜI ĐI: lợi nhuận = KẾT of that person (from RIGHT summary SUBTOTAL)
            // We reference the RIGHT summary totals which are at known positions
            // But it's simpler to just compute: SUMIFS(THU, NGƯỜI ĐI=nd) + ship + lấy + trả
            // Since RIGHT summary already calculates that, let's reference those cells.
            // However, to keep it simple and matching reference (hardcoded values),
            // we compute per-person totals directly from data.

            string nguoiDiColL = ColLetter(COL_NGUOIDI);
            // After column shift, RIGHT summary values are in L (COL_NGAYLAY)
            string rightValColL = ColLetter(COL_NGAYLAY);

            // Find RIGHT summary person blocks (they start at startRow)
            // Each person block is SUMMARY_RIGHT_BLOCK_HEIGHT+1 rows, total cell is at offset 6
            int rightBlockRow = startRow;
            foreach (string nd in distinctNguoiDis)
            {
                int totalCellRow = rightBlockRow + 6; // b6 in BuildRightSummary
                worksheet.Cell(curRow, COL_HANGTON).Value = nd.ToLower();
                // Reference the KẾT value from RIGHT summary
                worksheet.Cell(curRow, COL_FAIL).FormulaA1 = $"{rightValColL}{totalCellRow}";
                curRow++;
                rightBlockRow += SUMMARY_RIGHT_BLOCK_HEIGHT + 1;
            }

            // shop = -THANH TOÁN
            // Find THANH TOÁN cell: it's in the LEFT summary block at r8 position
            // LEFT starts at startRow, THANH TOÁN is at offset 8 within first block
            int thanhToanRow = startRow + 8; // r8 of first block
            worksheet.Cell(curRow, COL_HANGTON).Value = "shop";
            worksheet.Cell(curRow, COL_FAIL).FormulaA1 = $"-{tenkhColL}{thanhToanRow}";
            curRow++;

            // tổng = SUBTOTAL(9, all profit values)
            worksheet.Cell(curRow, COL_HANGTON).Value = "tổng";
            worksheet.Cell(curRow, COL_HANGTON).Style.Font.Bold = true;
            worksheet.Cell(curRow, COL_FAIL).FormulaA1 =
                $"SUBTOTAL(9,{ColLetter(COL_FAIL)}{firstValueRow}:{ColLetter(COL_FAIL)}{curRow - 1})";
            worksheet.Cell(curRow, COL_FAIL).Style.Font.Bold = true;
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
        /// Tra cứu phí ship theo QUẬN từ AppConstants.SHIPPING_FEES_BY_QUAN.
        /// Tự bỏ dấu tiếng Việt + normalize trước khi tra.
        /// </summary>
        private static decimal LookupShipFeeByQuan(string quan)
        {
            return LookupShipFeeByDict(quan, AppConstants.SHIPPING_FEES_BY_QUAN);
        }

        /// <summary>
        /// Tra cứu phí ship từ dictionary tùy ý (dùng cho cả bảng chung và bảng AT).
        /// Tự bỏ dấu + normalize quận trước khi tra. Trả về 0 nếu không tìm thấy.
        /// </summary>
        private static decimal LookupShipFeeByDict(string quan, Dictionary<string, decimal> feeDict)
        {
            if (string.IsNullOrWhiteSpace(quan))
                return 0m;
            if (feeDict.TryGetValue(quan.Trim(), out decimal fee1))
                return fee1;
            string norm = RemoveDiacriticsSimple(quan).ToLowerInvariant().Trim();
            norm = System.Text.RegularExpressions.Regex.Replace(
                norm,
                @"^(quan|huyen|tp|thanh pho)\s+",
                ""
            );
            if (feeDict.TryGetValue(norm, out decimal fee2))
                return fee2;
            var numMatch = System.Text.RegularExpressions.Regex.Match(norm, @"\d+");
            if (numMatch.Success && feeDict.TryGetValue(numMatch.Value, out decimal fee3))
                return fee3;
            return 0m;
        }

        private static string RemoveDiacriticsSimple(string s)
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

        /// <summary>
        /// Add header rows to new worksheet (row 1 = column headers, row 2 = THU x / NGAY x-x)
        /// </summary>
        private void AddHeaderRow(IXLWorksheet worksheet, DateTime date)
        {
            // Row 1: THU x | NGAY x-x (matches template format)
            string thuText;
            if (date.DayOfWeek == DayOfWeek.Sunday)
                thuText = "CHU NHAT";
            else
                thuText = "THU " + ((int)date.DayOfWeek + 1).ToString();

            string ngayText = "NGAY " + date.Day + "-" + date.Month;

            var cellThu = worksheet.Cell(1, COL_SHOP);
            cellThu.Value = thuText;
            cellThu.Style.Font.Bold = true;

            var cellNgay = worksheet.Cell(1, COL_TENKH);
            cellNgay.Value = ngayText;
            cellNgay.Style.Font.Bold = true;

            // Row 2: Column headers (16 columns matching template)
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
            };
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cell(2, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.LightGray;
            }
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

using System;
using System.IO;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace TextInputter.Services
{
    /// <summary>
    /// Service ghi dữ liệu invoice vào file Excel của khách.
    /// ⚠️ TODO: chưa được gọi từ UI — cần wire vào ExportMappedDataToExcel() trong OcrTab.cs.
    /// ⚠️ HARDCODED: tên file Excel mặc định phụ thuộc tháng — cần cập nhật mỗi tháng.
    /// </summary>
    public class ExcelInvoiceService
    {
        private readonly string _excelFilePath;

        /// <param name="excelFileName">
        /// ⚠️ HARDCODED tên file theo tháng — nhớ đổi khi sang tháng mới.
        /// </param>
        public ExcelInvoiceService(string excelFileName = "CHÂU NGÂN- THÁNG 2.2026- ĐỐI SOÁT.xlsx")
        {
            _excelFilePath = FindExcelFile(excelFileName);
        }

        /// <summary>
        /// Tìm file Excel theo tên trong các thư mục thông dụng.
        /// </summary>
        private string FindExcelFile(string fileName)
        {
            var searchPaths = new[]
            {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", fileName),
                Path.Combine(Directory.GetCurrentDirectory(), fileName)
            };

            foreach (var path in searchPaths)
            {
                if (File.Exists(path))
                    return path;
            }

            throw new FileNotFoundException($"Excel file not found: {fileName}");
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
                            if (row.RowNumber() <= 2) continue; // Skip header rows (row1=cols, row2=THU x)

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
                            if (row.RowNumber() <= 2) continue; // Skip header rows (row1=cols, row2=THU x)

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

        // Column index constants (1-based) matching existing sheet structure
        private const int COL_TINHTRANG  = 1;
        private const int COL_SHOP       = 2;
        private const int COL_TENKH      = 3;
        private const int COL_MA         = 4;
        private const int COL_SONHA      = 5;
        private const int COL_TENDUONG   = 6;
        private const int COL_QUAN       = 7;
        private const int COL_TIENTHU    = 8;
        private const int COL_TIENSHIP   = 9;
        private const int COL_TIENHANG   = 10;
        private const int COL_NGUOIDI    = 11;
        private const int COL_NGUOILAY   = 12;
        private const int COL_NGAYLAY    = 13;
        private const int COL_GHICHU     = 14;
        private const int COL_UNGIEN     = 15;
        private const int COL_HANGTON    = 16;
        private const int COL_FAIL       = 17;
        private const int COL_COL1       = 18;
        private const int COL_COL2       = 19;
        private const int COL_COL3       = 20;

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
                        if (row.RowNumber() <= 2) continue; // bỏ header rows
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
                        targetRow = (lastRow != null && lastRow.RowNumber() >= 3)
                            ? lastRow.RowNumber() + 1
                            : 3;
                    }

                    // Ghi dữ liệu vào targetRow
                    worksheet.Cell(targetRow, COL_SHOP).Value      = invoice.Shop ?? "";
                    worksheet.Cell(targetRow, COL_TENKH).Value     = invoice.TenKhachHang ?? "";
                    worksheet.Cell(targetRow, COL_MA).Value        = invoice.SoHoaDon;
                    worksheet.Cell(targetRow, COL_SONHA).Value     = invoice.SoNha;
                    worksheet.Cell(targetRow, COL_TENDUONG).Value  = invoice.TenDuong;
                    worksheet.Cell(targetRow, COL_QUAN).Value      = invoice.Quan;
                    worksheet.Cell(targetRow, COL_TIENTHU).Value   = invoice.TongThanhToan;
                    worksheet.Cell(targetRow, COL_TIENSHIP).Value  = 0;
                    worksheet.Cell(targetRow, COL_TIENHANG).Value  = invoice.TongTienHang;
                    worksheet.Cell(targetRow, COL_NGUOIDI).Value   = invoice.NguoiDi;
                    worksheet.Cell(targetRow, COL_NGUOILAY).Value  = invoice.NguoiLay;
                    worksheet.Cell(targetRow, COL_NGAYLAY).Value   = DateTime.Now.ToString("dd-MM-yyyy.");

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
                "Tình trạng TT", "SHOP", "TÊN KH", "MÃ", "SỐ NHÀ", "TÊN ĐƯỜNG", "QUẬN",
                "TIỀN THU", "TIỀN SHIP", "TIỀN HÀNG",
                "NGƯỜI ĐI", "NGƯỜI LẤY", "NGÀY LẤY", "GHI CHÚ",
                "ỨNG TIỀN", "HÀNG TỒN", "FAIL", "Column1", "Column2", "Column3"
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

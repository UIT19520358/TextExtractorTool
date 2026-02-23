using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using System.Data;
using System.Windows.Forms;

namespace TextInputter.Services
{
    /// <summary>
    /// Xử lý Excel operations
    /// </summary>
    public class ExcelService
    {
        /// <summary>
        /// Load Excel file và trả về tất cả sheets
        /// </summary>
        public Dictionary<string, DataTable> LoadExcelFile(string filePath)
        {
            var sheetData = new Dictionary<string, DataTable>();

            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        var dataTable = ConvertWorksheetToDataTable(worksheet);
                        sheetData[worksheet.Name] = dataTable;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi đọc Excel:\n{ex.Message}", "Lỗi");
            }

            return sheetData;
        }

        /// <summary>
        /// Convert worksheet thành DataTable
        /// </summary>
        private DataTable ConvertWorksheetToDataTable(IXLWorksheet worksheet)
        {
            var dataTable = new DataTable();
            var usedRange = worksheet.RangeUsed();

            if (usedRange == null)
                return dataTable;

            int rowCount = usedRange.RowCount();
            int colCount = usedRange.ColumnCount();

            // Tìm hàng header
            int headerRowIndex = 2;
            for (int row = 1; row <= Math.Min(5, rowCount); row++)
            {
                string firstCell = worksheet.Cell(row, 1).GetString()?.Trim() ?? "";
                if (firstCell == "SHOP" || firstCell.Contains("Tình trạng"))
                {
                    headerRowIndex = row;
                    break;
                }
            }

            // Add columns
            for (int col = 1; col <= colCount; col++)
            {
                string columnName = worksheet.Cell(headerRowIndex, col).GetString()?.Trim() ?? "";
                dataTable.Columns.Add(columnName);
            }

            // Add rows
            for (int row = 1; row <= rowCount; row++)
            {
                if (row == headerRowIndex)
                    continue;

                DataRow dataRow = dataTable.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    string cellValue = worksheet.Cell(row, col).GetString();
                    dataRow[col - 1] = cellValue ?? "";
                }
                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }

        /// <summary>
        /// Lưu Daily Report vào Excel với format đặc biệt
        /// </summary>
        public void SaveDailyReportToExcel(DataGridView dgv, decimal totalAmount, decimal totalDon, string excelPath)
        {
            try
            {
                string sheetName = DateTime.Now.ToString("dd-MM");
                XLWorkbook workbook;

                // Load existing or create new
                if (File.Exists(excelPath))
                {
                    workbook = new XLWorkbook(excelPath);
                    var existingSheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
                    if (existingSheet != null)
                    {
                        workbook.Worksheets.Delete(sheetName);
                    }
                }
                else
                {
                    workbook = new XLWorkbook();
                }

                using (workbook)
                {
                    var worksheet = workbook.Worksheets.Add(sheetName);

                    // Add headers
                    for (int col = 0; col < dgv.Columns.Count; col++)
                    {
                        worksheet.Cell(1, col + 1).Value = dgv.Columns[col].HeaderText;
                        worksheet.Cell(1, col + 1).Style.Font.Bold = true;
                        worksheet.Cell(1, col + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                    }

                    // Add data
                    for (int row = 0; row < dgv.Rows.Count; row++)
                    {
                        for (int col = 0; col < dgv.Columns.Count; col++)
                        {
                            var cellValue = dgv.Rows[row].Cells[col].Value;
                            worksheet.Cell(row + 2, col + 1).Value = cellValue?.ToString() ?? "";
                        }
                    }

                    // Add total row
                    int lastRow = dgv.Rows.Count + 2;
                    worksheet.Cell(lastRow, 1).Value = "TỔNG CỘNG";
                    worksheet.Cell(lastRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(lastRow, 1).Style.Fill.BackgroundColor = XLColor.Yellow;

                    worksheet.Cell(lastRow, 2).Value = totalAmount;
                    worksheet.Cell(lastRow, 2).Style.Font.Bold = true;
                    worksheet.Cell(lastRow, 2).Style.Fill.BackgroundColor = XLColor.LightBlue;

                    // Find and update SỐ ĐƠN column
                    for (int col = 0; col < dgv.Columns.Count; col++)
                    {
                        if (dgv.Columns[col].HeaderText.Contains("SỐ ĐƠN"))
                        {
                            worksheet.Cell(lastRow, col + 1).Value = totalDon;
                            worksheet.Cell(lastRow, col + 1).Style.Font.Bold = true;
                            worksheet.Cell(lastRow, col + 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
                            break;
                        }
                    }

                    worksheet.Columns().AdjustToContents();
                    workbook.SaveAs(excelPath);
                }

                MessageBox.Show($"✅ Lưu thành công vào: {excelPath}\n\nSheet: {sheetName}", "Thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi lưu: {ex.Message}", "Lỗi");
            }
        }
    }
}

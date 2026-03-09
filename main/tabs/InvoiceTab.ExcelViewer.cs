using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace TextInputter
{
    // ─── Excel Viewer + Loading Overlay ─────────────────────────────────────────
    public partial class MainForm
    {
        private void BtnOpenExcel_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter =
                        "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
                    openFileDialog.Title = "Chọn file Excel";
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                        LoadExcelFile(openFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌ Lỗi:\n{ex.Message}",
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private void LoadExcelFile(string filePath)
        {
            try
            {
                currentExcelFilePath = filePath;

                using (var workbook = new XLWorkbook(filePath))
                {
                    var sheetNames = workbook.Worksheets.Select(ws => ws.Name).ToList();
                    if (sheetNames.Count == 0)
                    {
                        MessageBox.Show("⚠️ File Excel không có sheet nào");
                        return;
                    }

                    tabExcelSheets.TabPages.Clear();

                    foreach (var sheetName in sheetNames)
                    {
                        TabPage tabPage = new TabPage(sheetName);
                        DataGridView dgv = new DataGridView
                        {
                            Dock = DockStyle.Fill,
                            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                            ReadOnly = false,
                            AllowUserToAddRows = true,
                            AllowUserToDeleteRows = true,
                        };
                        tabPage.Controls.Add(dgv);
                        LoadSheetData(workbook, sheetName, dgv);
                        tabExcelSheets.TabPages.Add(tabPage);
                    }

                    tabMainControl.SelectedTab = tabExcelViewer;
                    lblStatus.Text =
                        $"✅ Excel: {System.IO.Path.GetFileName(filePath)} ({sheetNames.Count} sheets)";
                    lblStatus.ForeColor = Color.Green;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi đọc Excel:\n{ex.Message}", "Lỗi");
                Debug.WriteLine($"Excel error: {ex.Message}");
            }
        }

        private void LoadSheetData(XLWorkbook workbook, string sheetName, DataGridView dgv)
        {
            try
            {
                var worksheet = workbook.Worksheet(sheetName);
                var usedRange = worksheet.RangeUsed();
                if (usedRange == null)
                    return;

                int rowCount = usedRange.RowCount();
                int colCount = usedRange.ColumnCount();

                // Tìm header row: quét tối đa 5 dòng đầu, tìm dòng chứa tên cột quen thuộc
                int headerRowIndex = -1;
                for (int row = 1; row <= Math.Min(5, rowCount); row++)
                {
                    // Quét tất cả cells trong row này để tìm keyword
                    bool rowHasHeader = false;
                    for (int col = 1; col <= Math.Min(colCount, 5); col++)
                    {
                        string cellVal = worksheet.Cell(row, col).GetString()?.Trim() ?? "";
                        if (
                            cellVal == "SHOP"
                            || cellVal.Contains("Tình trạng", StringComparison.OrdinalIgnoreCase)
                            || cellVal.Contains("TIỀN THU", StringComparison.OrdinalIgnoreCase)
                            || cellVal.Contains("TÊN KH", StringComparison.OrdinalIgnoreCase)
                        )
                        {
                            rowHasHeader = true;
                            break;
                        }
                    }
                    if (rowHasHeader)
                    {
                        headerRowIndex = row;
                        break;
                    }
                }
                // Nếu không tìm thấy header row → dùng cột mặc định theo vị trí (ExcelInvoiceService chuẩn)
                if (headerRowIndex < 0)
                {
                    // Không có header row trong file → fake DataTable với tên cột cố định
                    var fixedHeaders = new[]
                    {
                        "TÌNH TRẠNG TT",
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
                        "_col17",
                        "_col18",
                        "_col19",
                    };
                    System.Data.DataTable dtNoHeader = new System.Data.DataTable();
                    for (int col = 1; col <= colCount; col++)
                    {
                        string h =
                            (col - 1 < fixedHeaders.Length) ? fixedHeaders[col - 1] : $"_col{col}";
                        dtNoHeader.Columns.Add(h);
                    }
                    for (int row = 1; row <= rowCount; row++)
                    {
                        var dataRow = dtNoHeader.NewRow();
                        for (int col = 1; col <= colCount; col++)
                            dataRow[col - 1] = worksheet.Cell(row, col).GetString() ?? "";
                        dtNoHeader.Rows.Add(dataRow);
                    }
                    dgv.DataSource = dtNoHeader;
                    dgv.AutoResizeColumns();
                    return;
                }

                System.Data.DataTable dataTable = new System.Data.DataTable();
                for (int col = 1; col <= colCount; col++)
                {
                    string colHeader =
                        worksheet.Cell(headerRowIndex, col).GetString()?.Trim() ?? "";
                    // Nếu header trống → dùng tên placeholder để tránh DataTable duplicate/rename
                    if (string.IsNullOrEmpty(colHeader))
                        colHeader = $"_col{col}";
                    dataTable.Columns.Add(colHeader);
                }

                // Row ngay sau header là "THỨ x / NGÀY x-x" — bỏ qua, không phải đơn hàng
                int dayHeaderRowIndex = -1;
                if (headerRowIndex + 1 <= rowCount)
                {
                    string dayCell =
                        worksheet.Cell(headerRowIndex + 1, 2).GetString()?.Trim() ?? "";
                    if (
                        dayCell.StartsWith("THU ", StringComparison.OrdinalIgnoreCase)
                        || dayCell.StartsWith("THỨ ", StringComparison.OrdinalIgnoreCase)
                        || dayCell.Equals("CHU NHAT", StringComparison.OrdinalIgnoreCase)
                        || dayCell.Equals("CHỦ NHẬT", StringComparison.OrdinalIgnoreCase)
                    )
                        dayHeaderRowIndex = headerRowIndex + 1;
                }

                for (int row = 1; row <= rowCount; row++)
                {
                    if (row == headerRowIndex)
                        continue;
                    // Không skip row "THU x / NGAY x-x" nữa — đưa vào DGV để khi Lưu không bị mất
                    var dataRow = dataTable.NewRow();
                    for (int col = 1; col <= colCount; col++)
                        dataRow[col - 1] = worksheet.Cell(row, col).GetString() ?? "";
                    dataTable.Rows.Add(dataRow);
                }

                dgv.DataSource = dataTable;
                dgv.AutoResizeColumns();
                if (dgv.Rows.Count > 0)
                    dgv.Rows[0].Frozen = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Sheet error: {ex.Message}");
            }
        }

        // ─── Save / Undo / Cancel Excel Editor ────────────────────────────────

        private void BtnSaveExcelEditor_Click(object sender, EventArgs e)
        {
            try
            {
                if (tabExcelSheets.TabPages.Count == 0)
                {
                    MessageBox.Show("Chưa mở file Excel!");
                    return;
                }
                if (string.IsNullOrEmpty(currentExcelFilePath))
                {
                    MessageBox.Show("Không tìm thấy đường dẫn file Excel!", "Lỗi");
                    return;
                }

                using (var workbook = new XLWorkbook(currentExcelFilePath))
                {
                    foreach (TabPage tabPage in tabExcelSheets.TabPages)
                    {
                        var dgv = tabPage.Controls[0] as DataGridView;
                        if (dgv == null)
                            continue;

                        var worksheet = workbook.Worksheet(tabPage.Text);

                        // Tìm headerRowIndex giống LoadSheetData
                        var usedRange = worksheet.RangeUsed();
                        if (usedRange == null)
                            continue;
                        int rowCount = usedRange.RowCount();
                        int colCount = usedRange.ColumnCount();
                        int headerRow = -1;
                        for (int r = 1; r <= Math.Min(5, rowCount); r++)
                        {
                            bool found = false;
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
                                    found = true;
                                    break;
                                }
                            }
                            if (found)
                            {
                                headerRow = r;
                                break;
                            }
                        }
                        // Nếu không có header row trong file → ghi từ row 1 (không skip)
                        int dataStartRow = (headerRow > 0) ? headerRow + 1 : 1;
                        // Chỉ ghi lại data rows từ DGV vào đúng vị trí
                        // KHÔNG xóa sheet — giữ nguyên formatting, formulas
                        int excelRow = dataStartRow;
                        for (int dgvRow = 0; dgvRow < dgv.Rows.Count; dgvRow++)
                        {
                            if (dgv.Rows[dgvRow].IsNewRow)
                                continue;
                            for (int col = 0; col < dgv.Columns.Count && col < colCount; col++)
                            {
                                var cellValue = dgv.Rows[dgvRow].Cells[col].Value;
                                string strVal = cellValue?.ToString() ?? "";
                                var cell = worksheet.Cell(excelRow, col + 1);
                                // Chỉ ghi nếu cell không có formula (tránh xóa TIỀN HÀNG formula)
                                if (!cell.HasFormula)
                                {
                                    if (decimal.TryParse(strVal, out decimal numVal))
                                        cell.SetValue(numVal);
                                    else
                                        cell.SetValue(strVal);
                                }
                            }
                            excelRow++;
                        }
                    }
                    workbook.SaveAs(currentExcelFilePath);
                }

                MessageBox.Show("✅ Lưu file Excel thành công!", "Thành công");
                lblStatus.Text = $"✅ Lưu Excel: {System.IO.Path.GetFileName(currentExcelFilePath)}";
                lblStatus.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi lưu: {ex.Message}", "Lỗi");
                Debug.WriteLine($"Save Excel error: {ex.Message}");
            }
        }

        private void BtnUndoExcelEditor_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(currentExcelFilePath))
                {
                    LoadExcelFile(currentExcelFilePath);
                    MessageBox.Show("✅ Đã hoàn tác tất cả thay đổi!", "Thành công");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi: {ex.Message}", "Lỗi");
            }
        }

        private void BtnCancelExcelEditor_Click(object sender, EventArgs e)
        {
            try
            {
                tabExcelSheets.TabPages.Clear();
                currentExcelFilePath = "";
                lblStatus.Text = "✅ Đã đóng file Excel";
                lblStatus.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi: {ex.Message}", "Lỗi");
            }
        }

        // ─── Loading Overlay ───────────────────────────────────────────────────

        /// <summary>
        /// Hiện overlay loading (spinner + text) phủ lên toàn bộ form.
        /// Trả về Panel overlay để caller có thể Dispose khi xong.
        /// </summary>
        private Panel ShowLoadingOverlay(string message = "⏳ Đang xử lý...")
        {
            var overlay = new Panel
            {
                BackColor = Color.FromArgb(160, 0, 0, 0),
                Dock = DockStyle.Fill,
                Name = "_loadingOverlay",
            };

            var lbl = new Label
            {
                Text = message,
                ForeColor = Color.White,
                Font = new Font("Arial", 13, FontStyle.Bold),
                AutoSize = false,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
            };

            var spinner = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30,
                Height = 8,
                Dock = DockStyle.Bottom,
            };

            overlay.Controls.Add(lbl);
            overlay.Controls.Add(spinner);

            this.Controls.Add(overlay);
            overlay.BringToFront();
            overlay.Refresh();
            return overlay;
        }

        /// <summary>Ẩn và dispose overlay loading.</summary>
        private void HideLoadingOverlay(Panel overlay)
        {
            if (overlay == null)
                return;
            this.Controls.Remove(overlay);
            overlay.Dispose();
        }

        // ─── IsDateLabelRow helper ─────────────────────────────────────────────

        /// <summary>
        /// Nhận biết dòng label ngày tháng (VD: SHOP="THU 2", TÊN KH="NGAY 02-03").
        /// Những dòng này không phải đơn hàng thực sự → bỏ qua khi tính tiền.
        /// </summary>
        private static bool IsDateLabelRow(DataGridViewRow row, int colShop, int colMa)
        {
            if (colShop < 0 || colShop >= row.Cells.Count)
                return false;
            string shop = row.Cells[colShop].Value?.ToString()?.Trim() ?? "";
            // "THU 2" .. "THU 7" — không có MÃ HĐ
            if (
                !System.Text.RegularExpressions.Regex.IsMatch(
                    shop,
                    @"^THU\s*\d$",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                )
            )
                return false;
            // Chắc chắn hơn: cột MÃ phải rỗng
            if (
                colMa >= 0
                && colMa < row.Cells.Count
                && !string.IsNullOrWhiteSpace(row.Cells[colMa].Value?.ToString())
            )
                return false;
            return true;
        }
    }
}

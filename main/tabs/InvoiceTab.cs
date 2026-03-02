using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace TextInputter
{
    /// <summary>
    /// Invoice / Excel Viewer Tab: mở Excel, hiển thị DataGridView, tính toán daily report
    /// </summary>
    public partial class MainForm
    {
        // ─── Helper class ──────────────────────────────────────────────────────

        private class DailyReportData
        {
            public string Date { get; set; }
            public decimal TongTienThu { get; set; } // Tổng tiền thu (cột H)
            public decimal TongTienShip { get; set; } // Tổng tiền ship (cột I)
            public decimal KhoanTruShip { get; set; } // -(TongShip - SoDon×5), số âm
            public decimal TongKetCuoi { get; set; } // TongTienThu + KhoanTruShip
            public decimal SoDon { get; set; }

            // Các row âm (đơn trả, đơn cũ ck...) lấy từ Excel
            public List<(string Label, decimal Amount)> NegativeRows { get; set; } = new();

            // Report nhỏ theo từng người đi: Key = tên người, Value = (TienThu, TienShip, SoDon)
            public Dictionary<
                string,
                (decimal TienThu, decimal TienShip, decimal SoDon)
            > ReportByNguoiDi { get; set; } =
                new Dictionary<string, (decimal, decimal, decimal)>(
                    StringComparer.OrdinalIgnoreCase
                );
        }

        private DailyReportData currentDailyReport;

        // ─── Excel Viewer ──────────────────────────────────────────────────────

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

                System.Data.DataTable dataTable = new System.Data.DataTable();
                for (int col = 1; col <= colCount; col++)
                    dataTable.Columns.Add(
                        worksheet.Cell(headerRowIndex, col).GetString()?.Trim() ?? ""
                    );

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
                    if (row == dayHeaderRowIndex)
                        continue; // bỏ qua row "THỨ x | NGÀY x-x"
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
                        worksheet.Clear();

                        for (int col = 0; col < dgv.Columns.Count; col++)
                            worksheet.Cell(1, col + 1).Value = dgv.Columns[col].HeaderText;

                        for (int row = 0; row < dgv.Rows.Count; row++)
                        for (int col = 0; col < dgv.Columns.Count; col++)
                        {
                            var cellValue = dgv.Rows[row].Cells[col].Value;
                            if (cellValue != null)
                                worksheet.Cell(row + 2, col + 1).Value = cellValue.ToString();
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

        // ─── Calculate (Excel → Daily Report) ─────────────────────────────────

        private void BtnCalculateExcelData_Click(object sender, EventArgs e)
        {
            try
            {
                if (tabExcelSheets.TabPages.Count == 0)
                    return;

                var currentSheet = tabExcelSheets.SelectedTab;
                if (currentSheet == null || currentSheet.Controls.Count == 0)
                    return;

                DataGridView sourceGridView = null;
                foreach (Control ctrl in currentSheet.Controls)
                    if (ctrl is DataGridView dgv)
                    {
                        sourceGridView = dgv;
                        break;
                    }

                if (sourceGridView == null || sourceGridView.Rows.Count == 0)
                    return;

                // Column detection
                int colShop = -1,
                    colTienThu = -1,
                    colTienShip = -1,
                    colTienHang = -1,
                    colSoDon = -1,
                    colGhiChu = -1,
                    colNgayLay = -1,
                    colNguoiDi = -1;
                for (int col = 0; col < sourceGridView.Columns.Count; col++)
                {
                    string header = sourceGridView.Columns[col].HeaderText.ToLower();
                    if (header.Contains("shop"))
                        colShop = col;
                    if (header.Contains("tiền thu"))
                        colTienThu = col;
                    if (header.Contains("tiền ship"))
                        colTienShip = col;
                    if (header.Contains("tiền hàng"))
                        colTienHang = col;
                    if (header.Contains("số đơn"))
                        colSoDon = col;
                    if (header.Contains("ghi chú"))
                        colGhiChu = col;
                    if (header.Contains("ngày lấy"))
                        colNgayLay = col;
                    if (header.Contains("người đi") || header.Contains("nguoi di"))
                        colNguoiDi = col;
                }

                Debug.WriteLine(
                    $"Cols — Shop:{colShop} TienThu:{colTienThu} TienShip:{colTienShip} TienHang:{colTienHang} SoDon:{colSoDon}"
                );

                // PHẦN 1: Copy dữ liệu sang dgvInvoice
                dgvInvoice.DataSource = null;
                dgvInvoice.Rows.Clear();
                dgvInvoice.Columns.Clear();

                foreach (DataGridViewColumn col in sourceGridView.Columns)
                    dgvInvoice.Columns.Add(col.Name, col.HeaderText);

                // Tìm colTienHang sớm để lọc row âm
                int colTienHangCheck = colTienHang;

                // Tìm colMa một lần
                int colMa = -1;
                for (int c = 0; c < sourceGridView.Columns.Count; c++)
                    if (sourceGridView.Columns[c].HeaderText.ToLower().Contains("mã"))
                    {
                        colMa = c;
                        break;
                    }

                // ── BƯỚC 1: Tìm SUM row trong Excel ────────────────────────────────
                decimal totalTienThu = 0,
                    totalTienShip = 0,
                    totalSoDon = 0;
                bool foundSumRow = false;
                int sumRowIndex = -1;

                for (int i = 0; i < sourceGridView.Rows.Count; i++)
                {
                    var row = sourceGridView.Rows[i];
                    if (row.IsNewRow)
                        continue;
                    string shopVal = colShop >= 0 ? row.Cells[colShop].Value?.ToString() ?? "" : "";
                    if (!string.IsNullOrWhiteSpace(shopVal))
                        continue;

                    int checkCol = colTienThu >= 0 ? colTienThu : colTienHang;
                    if (checkCol < 0 || checkCol >= row.Cells.Count)
                        continue;
                    if (
                        !decimal.TryParse(
                            row.Cells[checkCol].Value?.ToString() ?? "",
                            out decimal chkVal
                        )
                        || chkVal <= 0
                    )
                        continue;

                    sumRowIndex = i;
                    foundSumRow = true;
                    if (colTienThu >= 0)
                        decimal.TryParse(row.Cells[colTienThu].Value?.ToString(), out totalTienThu);
                    if (colTienShip >= 0)
                        decimal.TryParse(
                            row.Cells[colTienShip].Value?.ToString(),
                            out totalTienShip
                        );
                    if (colSoDon >= 0)
                        decimal.TryParse(row.Cells[colSoDon].Value?.ToString(), out totalSoDon);
                    // Fallback: cột Column1 chứa SỐ ĐƠN khi header không detect được
                    if (totalSoDon == 0 && row.Cells.Count > AppConstants.COL_SODON_FALLBACK_IDX)
                        decimal.TryParse(
                            row.Cells[AppConstants.COL_SODON_FALLBACK_IDX].Value?.ToString(),
                            out totalSoDon
                        );
                    // Log toàn bộ cells của SUM row để debug
                    var sbDebug = new System.Text.StringBuilder();
                    for (int dc = 0; dc < row.Cells.Count; dc++)
                        sbDebug.Append($"[{dc}]={row.Cells[dc].Value} ");
                    Debug.WriteLine($"SUM row idx={i}: {sbDebug}");
                    Debug.WriteLine(
                        $"SUM row idx={i}: TienThu={totalTienThu}, Ship={totalTienShip}, SoDon={totalSoDon}"
                    );
                    break;
                }

                // Nếu không có SUM row → tự cộng từng row DATA
                if (!foundSumRow)
                {
                    foreach (DataGridViewRow row in sourceGridView.Rows)
                    {
                        if (row.IsNewRow)
                            continue;
                        string sv = colShop >= 0 ? row.Cells[colShop].Value?.ToString() ?? "" : "";
                        if (string.IsNullOrWhiteSpace(sv))
                            continue;
                        if (colTienThu >= 0)
                        {
                            if (
                                decimal.TryParse(
                                    row.Cells[colTienThu].Value?.ToString(),
                                    out decimal t
                                )
                            )
                                totalTienThu += t;
                        }
                        if (colTienShip >= 0)
                        {
                            if (
                                decimal.TryParse(
                                    row.Cells[colTienShip].Value?.ToString(),
                                    out decimal s
                                )
                            )
                                totalTienShip += s;
                        }
                        totalSoDon++;
                    }
                }

                // Thu thập các row âm (đơn trả, đơn cũ ck):
                // Điều kiện nhận dạng "row âm khoản trừ" (phân biệt với đơn có MÃ mà TIỀN HÀNG âm):
                //   • TIỀN HÀNG < 0  (bắt buộc)
                //   • KHÔNG có MÃ HĐ (colMa rỗng/null)  ← đơn thật sẽ luôn có mã
                //   • KHÔNG có SHOP  (colShop rỗng/null) ← đơn thật sẽ luôn có shop
                // Nếu có SUM row → chỉ tìm SAU SUM row.
                // Nếu không có SUM row → quét toàn bộ nhưng vẫn giữ điều kiện lọc trên.
                var negativeRows = new List<DataGridViewRow>();
                if (colTienHangCheck >= 0)
                {
                    int startIdx = foundSumRow ? sumRowIndex + 1 : 0;
                    for (int i = startIdx; i < sourceGridView.Rows.Count; i++)
                    {
                        var row = sourceGridView.Rows[i];
                        if (row.IsNewRow)
                            continue;
                        if (colTienHangCheck >= row.Cells.Count)
                            continue;
                        if (
                            !decimal.TryParse(
                                row.Cells[colTienHangCheck].Value?.ToString() ?? "",
                                out decimal jVal
                            )
                            || jVal >= 0
                        )
                            continue;

                        // Loại bỏ nếu có MÃ HĐ (đơn thật bị âm, không phải khoản trừ)
                        if (
                            colMa >= 0
                            && colMa < row.Cells.Count
                            && !string.IsNullOrWhiteSpace(row.Cells[colMa].Value?.ToString())
                        )
                            continue;
                        // Loại bỏ nếu có SHOP (đơn thật bị âm, không phải khoản trừ)
                        if (
                            colShop >= 0
                            && colShop < row.Cells.Count
                            && !string.IsNullOrWhiteSpace(row.Cells[colShop].Value?.ToString())
                        )
                            continue;

                        negativeRows.Add(row);
                    }
                }

                // Tính tổng số âm ở TIỀN HÀNG
                decimal totalNegHang = 0;
                foreach (var nr in negativeRows)
                    if (
                        decimal.TryParse(
                            nr.Cells[colTienHangCheck].Value?.ToString() ?? "",
                            out decimal nv
                        )
                    )
                        totalNegHang += nv;

                decimal tongHangDuong = totalTienThu - totalTienShip; // SUM row TIỀN HÀNG
                decimal tongKetCuoi = tongHangDuong + totalNegHang; // cộng luôn số âm
                decimal phiShipThucTe = totalSoDon * AppConstants.PHI_SHIP_MOI_DON;
                decimal khoanTruShip = -(totalTienShip - phiShipThucTe);

                // ── Tổng hợp theo NGƯỜI ĐI ──────────────────────────────────────
                // Quét toàn bộ data rows (trước SUM row), gom tiền thu + tiền ship + số đơn theo người đi.
                var reportByNguoiDi = new Dictionary<
                    string,
                    (decimal TienThu, decimal TienShip, decimal SoDon)
                >(StringComparer.OrdinalIgnoreCase);
                if (colNguoiDi >= 0)
                {
                    int endIdx = sumRowIndex >= 0 ? sumRowIndex : sourceGridView.Rows.Count;
                    for (int i = 0; i < endIdx; i++)
                    {
                        var row = sourceGridView.Rows[i];
                        if (row.IsNewRow)
                            continue;

                        // Chỉ lấy data rows (có SHOP)
                        string sv = colShop >= 0 ? row.Cells[colShop].Value?.ToString() ?? "" : "";
                        if (string.IsNullOrWhiteSpace(sv))
                            continue;

                        string nguoiRow =
                            colNguoiDi < row.Cells.Count
                                ? (row.Cells[colNguoiDi].Value?.ToString() ?? "").Trim()
                                : "";
                        if (string.IsNullOrEmpty(nguoiRow))
                            nguoiRow = "(không rõ)";

                        decimal tienThuRow = 0;
                        if (colTienThu >= 0 && colTienThu < row.Cells.Count)
                            decimal.TryParse(
                                row.Cells[colTienThu].Value?.ToString() ?? "",
                                out tienThuRow
                            );

                        decimal tienShipRow = 0;
                        if (colTienShip >= 0 && colTienShip < row.Cells.Count)
                            decimal.TryParse(
                                row.Cells[colTienShip].Value?.ToString() ?? "",
                                out tienShipRow
                            );

                        if (!reportByNguoiDi.ContainsKey(nguoiRow))
                            reportByNguoiDi[nguoiRow] = (0, 0, 0);
                        var cur = reportByNguoiDi[nguoiRow];
                        reportByNguoiDi[nguoiRow] = (
                            cur.TienThu + tienThuRow,
                            cur.TienShip + tienShipRow,
                            cur.SoDon + 1
                        );
                    }
                }

                Debug.WriteLine(
                    $"FINAL: SumRow={foundSumRow}, Thu={totalTienThu}, Ship={totalTienShip}, HangDuong={tongHangDuong}, NegHang={totalNegHang}, KetCuoi={tongKetCuoi}"
                );

                // ── BƯỚC 2: Build dgvInvoice đúng thứ tự ───────────────────────────
                dgvInvoice.DataSource = null;
                dgvInvoice.Rows.Clear();
                dgvInvoice.Columns.Clear();
                foreach (DataGridViewColumn col in sourceGridView.Columns)
                    dgvInvoice.Columns.Add(col.Name, col.HeaderText);

                void AddRow(DataGridViewRow src, Color? bg, bool italic)
                {
                    var r = new DataGridViewRow();
                    r.CreateCells(dgvInvoice);
                    for (int ci = 0; ci < src.Cells.Count && ci < r.Cells.Count; ci++)
                        r.Cells[ci].Value = src.Cells[ci].Value;
                    dgvInvoice.Rows.Add(r);
                    int idx = dgvInvoice.Rows.Count - 1;
                    if (bg.HasValue)
                        for (int ci = 0; ci < dgvInvoice.Columns.Count; ci++)
                            dgvInvoice.Rows[idx].Cells[ci].Style.BackColor = bg.Value;
                    if (italic)
                        for (int ci = 0; ci < dgvInvoice.Columns.Count; ci++)
                            dgvInvoice.Rows[idx].Cells[ci].Style.Font = new Font(
                                dgvInvoice.Font,
                                FontStyle.Italic
                            );
                }

                // 1. Data rows (có SHOP, bao gồm cả đơn không có MÃ)
                for (
                    int i = 0;
                    i < (sumRowIndex >= 0 ? sumRowIndex : sourceGridView.Rows.Count);
                    i++
                )
                {
                    var sr = sourceGridView.Rows[i];
                    if (sr.IsNewRow)
                        continue;
                    string sv = colShop >= 0 ? sr.Cells[colShop].Value?.ToString() ?? "" : "";
                    if (string.IsNullOrWhiteSpace(sv))
                        continue;
                    AddRow(sr, null, false);
                }

                // 2. SUM row — màu vàng
                {
                    var sumRow = new DataGridViewRow();
                    sumRow.CreateCells(dgvInvoice);
                    if (sumRow.Cells.Count > 0)
                        sumRow.Cells[0].Value = "▶ TỔNG";
                    if (colTienThu >= 0 && colTienThu < sumRow.Cells.Count)
                        sumRow.Cells[colTienThu].Value = totalTienThu.ToString();
                    if (colTienShip >= 0 && colTienShip < sumRow.Cells.Count)
                        sumRow.Cells[colTienShip].Value = totalTienShip.ToString();
                    if (colTienHang >= 0 && colTienHang < sumRow.Cells.Count)
                        sumRow.Cells[colTienHang].Value = tongHangDuong.ToString();
                    if (colSoDon >= 0 && colSoDon < sumRow.Cells.Count)
                        sumRow.Cells[colSoDon].Value = totalSoDon.ToString();
                    // Không ghi fallback vào cells[16] vì sẽ đè vào cột FAIL
                    dgvInvoice.Rows.Add(sumRow);
                    int si = dgvInvoice.Rows.Count - 1;
                    for (int ci = 0; ci < dgvInvoice.Columns.Count; ci++)
                    {
                        dgvInvoice.Rows[si].Cells[ci].Style.BackColor = AppConstants.COLOR_ROW_TONG;
                        dgvInvoice.Rows[si].Cells[ci].Style.ForeColor = Color.Black;
                        dgvInvoice.Rows[si].Cells[ci].Style.Font = new Font(
                            dgvInvoice.Font,
                            FontStyle.Bold
                        );
                    }
                    dgvInvoice.Rows[si].Height = AppConstants.ROW_HEIGHT_TONG;
                }

                // 3. Row âm — màu cam italic (giữ nguyên từ Excel)
                foreach (var nr in negativeRows)
                    AddRow(nr, AppConstants.COLOR_ROW_NEGATIVE, true);

                // 4. Dòng KẾT cuối = SUM + số âm — chỉ hiện khi có row âm
                if (negativeRows.Count > 0)
                {
                    var ketRow = new DataGridViewRow();
                    ketRow.CreateCells(dgvInvoice);
                    if (ketRow.Cells.Count > 0)
                        ketRow.Cells[0].Value = "▶ KẾT";
                    if (colTienHang >= 0 && colTienHang < ketRow.Cells.Count)
                        ketRow.Cells[colTienHang].Value = tongKetCuoi.ToString();
                    if (colSoDon >= 0 && colSoDon < ketRow.Cells.Count)
                        ketRow.Cells[colSoDon].Value = totalSoDon.ToString();
                    // Fallback cột fallback index nếu không detect colSoDon
                    if (colSoDon < 0 && ketRow.Cells.Count > AppConstants.COL_SODON_FALLBACK_IDX)
                        ketRow.Cells[AppConstants.COL_SODON_FALLBACK_IDX].Value =
                            totalSoDon.ToString();
                    dgvInvoice.Rows.Add(ketRow);
                    int ki = dgvInvoice.Rows.Count - 1;
                    for (int ci = 0; ci < dgvInvoice.Columns.Count; ci++)
                    {
                        dgvInvoice.Rows[ki].Cells[ci].Style.BackColor = AppConstants.COLOR_ROW_KET;
                        dgvInvoice.Rows[ki].Cells[ci].Style.ForeColor = Color.Black;
                        dgvInvoice.Rows[ki].Cells[ci].Style.Font = new Font(
                            dgvInvoice.Font,
                            FontStyle.Bold
                        );
                    }
                    dgvInvoice.Rows[ki].Height = AppConstants.ROW_HEIGHT_KET;
                }

                // Lấy ngày lấy từ data (dùng làm sheet name khi Save)
                string reportDate = DateTime.Now.ToString("dd-MM-yyyy"); // fallback
                if (colNgayLay >= 0)
                {
                    foreach (DataGridViewRow dr in sourceGridView.Rows)
                    {
                        string ngay = dr.Cells[colNgayLay].Value?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(ngay))
                        {
                            // Normalize: bỏ dấu chấm/gạch chéo, đổi sang dd-MM-yyyy
                            if (DateTime.TryParse(ngay, out DateTime dt))
                                reportDate = dt.ToString("dd-MM-yyyy");
                            else
                                reportDate = ngay.Replace("/", "-").Replace(".", "-");
                            break;
                        }
                    }
                }

                currentDailyReport = new DailyReportData
                {
                    Date = reportDate,
                    TongTienThu = totalTienThu,
                    TongTienShip = totalTienShip,
                    KhoanTruShip = khoanTruShip,
                    TongKetCuoi = tongKetCuoi,
                    SoDon = totalSoDon,
                    ReportByNguoiDi = reportByNguoiDi,
                    NegativeRows = negativeRows
                        .Select(nr =>
                        {
                            // Tìm label: quét tất cả cells, lấy ô có text (không phải số, không rỗng)
                            string lbl = "";
                            for (int ci = 0; ci < nr.Cells.Count; ci++)
                            {
                                string v = nr.Cells[ci].Value?.ToString()?.Trim() ?? "";
                                if (string.IsNullOrEmpty(v))
                                    continue;
                                if (decimal.TryParse(v, out _))
                                    continue; // bỏ qua ô số
                                lbl = v;
                                break;
                            }
                            if (string.IsNullOrEmpty(lbl))
                                lbl = "đơn âm";
                            decimal.TryParse(
                                nr.Cells[colTienHangCheck].Value?.ToString() ?? "",
                                out decimal amt
                            );
                            return (lbl, amt);
                        })
                        .ToList(),
                };

                lblInvoiceTotal.Text =
                    $"TỔNG THU: {totalTienThu:N0} đ | SHIP: {totalTienShip:N0} đ | SỐ ĐƠN: {totalSoDon:N0} | KẾT: {tongKetCuoi:N0} đ";

                DisplayDailyReport();
                InitializeInvoiceButtonPanel();
                tabMainControl.SelectedIndex = 2;

                lblStatus.Text = "✅ Đã tính tiền — bấm 💾 Lưu để ghi vào Excel";
                lblStatus.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ Lỗi: {ex.Message}");
            }
        }

        // ─── Invoice dgv helpers ───────────────────────────────────────────────

        private void BtnAddInvoiceRow_Click(object sender, EventArgs e)
        {
            if (dgvInvoice.Columns.Count == 0)
            {
                dgvInvoice.Columns.Add("Tên", "Tên");
                dgvInvoice.Columns.Add("Tiền", "Tiền");
                dgvInvoice.Columns.Add("Số đơn", "Số đơn");
            }
            dgvInvoice.Rows.Add("", "0", "0");
        }

        private void BtnCalculateInvoice_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvInvoice.Rows.Count == 0)
                {
                    MessageBox.Show("Chưa có dữ liệu để tính!");
                    return;
                }

                decimal totalTien = 0,
                    totalSoDon = 0;
                for (int i = 0; i < dgvInvoice.Rows.Count; i++)
                {
                    if (
                        decimal.TryParse(
                            dgvInvoice.Rows[i].Cells[1].Value?.ToString() ?? "0",
                            out decimal tienHang
                        )
                    )
                        totalTien += tienHang;
                    if (
                        decimal.TryParse(
                            dgvInvoice.Rows[i].Cells.Count > 8
                                ? dgvInvoice.Rows[i].Cells[8].Value?.ToString() ?? "0"
                                : "0",
                            out decimal sodon
                        )
                    )
                        totalSoDon += sodon;
                }

                lblInvoiceTotal.Text = $"TỔNG CỘNG: {totalTien:N0} đ | SỐ ĐƠN: {totalSoDon:N0}";

                currentDailyReport = new DailyReportData
                {
                    Date = DateTime.Now.ToString("dd.MM.yyyy"),
                    TongTienThu = totalTien,
                    TongTienShip = 0,
                    KhoanTruShip = 0,
                    TongKetCuoi = totalTien,
                    SoDon = totalSoDon,
                };

                InitializeInvoiceButtonPanel();
                DisplayDailyReport();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ Lỗi: {ex.Message}");
            }
        }

        // ─── Daily Report Display ──────────────────────────────────────────────

        private void DisplayDailyReport()
        {
            if (currentDailyReport == null)
                return;

            Panel pnlTop = tabInvoice.Controls["pnlInvoiceTop"] as Panel;
            Panel pnlBottom = tabInvoice.Controls["pnlDailyReportBottom"] as Panel;

            if (pnlTop == null)
            {
                tabInvoice.Controls.Clear();

                pnlTop = new Panel
                {
                    Name = "pnlInvoiceTop",
                    Dock = DockStyle.Fill,
                    BackColor = Color.White,
                };
                pnlTop.Controls.Add(dgvInvoice);
                pnlTop.Controls.Add(lblInvoiceTotal);
                tabInvoice.Controls.Add(pnlTop);

                pnlBottom = new Panel
                {
                    Name = "pnlDailyReportBottom",
                    Dock = DockStyle.Bottom,
                    BackColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    Height = AppConstants.DAILY_REPORT_PANEL_HEIGHT,
                };
                tabInvoice.Controls.Add(pnlBottom);
            }

            pnlBottom.Controls.Clear();

            var r = currentDailyReport;
            string soDonStr = r.SoDon.ToString("N0");
            string thuStr = r.TongTienThu.ToString("N0");
            decimal tongShipRaw = -r.TongTienShip; // -SUMIFS toàn bộ TIỀN SHIP
            decimal tienLayTong = -(r.SoDon * AppConstants.PHI_SHIP_MOI_DON); // -(số đơn × 5)
            // KẾT = TongThu + tiền ship (âm) + tiền lấy (âm)
            decimal ketTong = r.TongTienThu + tongShipRaw + tienLayTong;
            string ketStr = ketTong.ToString("N0");

            Debug.WriteLine(
                $"DisplayDailyReport: TongThu={r.TongTienThu}, TongShip={r.TongTienShip}, KhoanTru={r.KhoanTruShip}, TongKet={r.TongKetCuoi}, SoDon={r.SoDon}"
            );

            // ── Helper: tạo 1 DataGridView report nhỏ ─────────────────────────
            DataGridView MakeReportGrid()
            {
                var g = new DataGridView
                {
                    BackgroundColor = Color.White,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    ReadOnly = true,
                    ColumnHeadersVisible = false,
                    RowHeadersVisible = false,
                    ScrollBars = ScrollBars.Vertical,
                    DefaultCellStyle =
                    {
                        Font = new Font("Arial", 10),
                        Alignment = DataGridViewContentAlignment.MiddleLeft,
                    },
                    AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                };
                g.Columns.Add("TenMuc", "");
                g.Columns.Add("Tien", "");
                g.Columns.Add("SoDon", "");
                g.Columns[0].Width = 220;
                g.Columns[1].Width = 110;
                g.Columns[2].Width = 90;
                g.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                g.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                return g;
            }

            // ── Panel chứa tất cả reports theo chiều ngang ────────────────────
            // Layout: [Report Tổng] | [Report người 1] | [Report người 2] | ...
            var pnlReports = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.White,
            };
            pnlBottom.Controls.Add(pnlReports);

            int panelWidth = 450;
            int panelX = 0;

            // ── Report TỔNG (bên trái) ────────────────────────────────────────
            {
                var pnlTong = new Panel
                {
                    Location = new Point(panelX, 0),
                    Width = panelWidth,
                    Height = pnlBottom.Height - 4,
                    BorderStyle = BorderStyle.FixedSingle,
                    BackColor = Color.White,
                };
                panelX += panelWidth + 6;

                var lblTong = new Label
                {
                    Text = "📊 TỔNG HỢP",
                    Dock = DockStyle.Top,
                    Height = 22,
                    Font = new Font("Arial", 9, FontStyle.Bold),
                    BackColor = Color.LightSteelBlue,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                };
                pnlTong.Controls.Add(lblTong);

                var dgvTong = MakeReportGrid();
                dgvTong.Dock = DockStyle.Fill;

                int ri;
                ri = dgvTong.Rows.Add("", "Tiền Thu", "Số đơn");
                dgvTong.Rows[ri].DefaultCellStyle.BackColor = Color.LightSteelBlue;
                dgvTong.Rows[ri].DefaultCellStyle.Font = new Font("Arial", 10, FontStyle.Bold);

                ri = dgvTong.Rows.Add("TỔNG ĐƠN", thuStr, soDonStr);
                dgvTong.Rows[ri].DefaultCellStyle.BackColor = Color.White;

                ri = dgvTong.Rows.Add("tiền ship", tongShipRaw.ToString("N0"), "");
                dgvTong.Rows[ri].DefaultCellStyle.BackColor = Color.White;
                dgvTong.Rows[ri].Cells[1].Style.ForeColor =
                    tongShipRaw < 0 ? Color.Red : Color.Black;

                ri = dgvTong.Rows.Add("tiền lấy", tienLayTong.ToString("N0"), "");
                dgvTong.Rows[ri].DefaultCellStyle.BackColor = Color.White;
                dgvTong.Rows[ri].Cells[1].Style.ForeColor =
                    tienLayTong < 0 ? Color.Red : Color.Black;

                ri = dgvTong.Rows.Add("đơn trả", "", "");
                dgvTong.Rows[ri].DefaultCellStyle.ForeColor = Color.Red;

                ri = dgvTong.Rows.Add("đơn cũ ck", "", "");
                dgvTong.Rows[ri].DefaultCellStyle.ForeColor = Color.Red;

                ri = dgvTong.Rows.Add("", ketStr, soDonStr);
                dgvTong.Rows[ri].DefaultCellStyle.BackColor = AppConstants.COLOR_REPORT_KET;
                dgvTong.Rows[ri].DefaultCellStyle.Font = new Font("Arial", 11, FontStyle.Bold);
                dgvTong.Rows[ri].Height = AppConstants.ROW_HEIGHT_REPORT_KET;

                pnlTong.Controls.Add(dgvTong);
                pnlReports.Controls.Add(pnlTong);
            }

            // ── Report nhỏ theo từng NGƯỜI ĐI ────────────────────────────────
            if (r.ReportByNguoiDi != null && r.ReportByNguoiDi.Count > 0)
            {
                int nguoiPanelWidth = 340;
                foreach (var kvp in r.ReportByNguoiDi.OrderBy(k => k.Key))
                {
                    string tenNguoi = kvp.Key;
                    decimal tienThuNguoi = kvp.Value.TienThu;
                    decimal tienShipNguoi = kvp.Value.TienShip;
                    decimal soDonNguoi = kvp.Value.SoDon;

                    var pnlNguoi = new Panel
                    {
                        Location = new Point(panelX, 0),
                        Width = nguoiPanelWidth,
                        Height = pnlBottom.Height - 4,
                        BorderStyle = BorderStyle.FixedSingle,
                        BackColor = Color.White,
                    };
                    panelX += nguoiPanelWidth + 6;

                    var lblNguoi = new Label
                    {
                        Text = $"👤 {tenNguoi.ToUpper()}",
                        Dock = DockStyle.Top,
                        Height = 22,
                        Font = new Font("Arial", 9, FontStyle.Bold),
                        BackColor = Color.FromArgb(200, 230, 255),
                        TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                    };
                    pnlNguoi.Controls.Add(lblNguoi);

                    var dgvNguoi = MakeReportGrid();
                    dgvNguoi.Dock = DockStyle.Fill;
                    dgvNguoi.Columns[0].Width = 150;
                    dgvNguoi.Columns[1].Width = 100;
                    dgvNguoi.Columns[2].Width = 70;

                    int ri;
                    // Header
                    ri = dgvNguoi.Rows.Add("", "Tiền Thu", "Số đơn");
                    dgvNguoi.Rows[ri].DefaultCellStyle.BackColor = Color.FromArgb(200, 230, 255);
                    dgvNguoi.Rows[ri].DefaultCellStyle.Font = new Font("Arial", 10, FontStyle.Bold);

                    // TỔNG ĐƠN NHẬN
                    ri = dgvNguoi.Rows.Add(
                        "TỔNG ĐƠN",
                        tienThuNguoi.ToString("N0"),
                        soDonNguoi.ToString("N0")
                    );
                    dgvNguoi.Rows[ri].DefaultCellStyle.BackColor = Color.White;

                    // tiền ship = -(tổng tiền ship của người đó)
                    decimal khoanShipNguoi = -tienShipNguoi;
                    ri = dgvNguoi.Rows.Add("tiền ship", khoanShipNguoi.ToString("N0"), "");
                    dgvNguoi.Rows[ri].DefaultCellStyle.BackColor = Color.White;
                    dgvNguoi.Rows[ri].Cells[1].Style.ForeColor =
                        khoanShipNguoi < 0 ? Color.Red : Color.Black;

                    // tiền lấy = -(số đơn × 5)
                    decimal tienLayNguoi = -(soDonNguoi * AppConstants.PHI_SHIP_MOI_DON);
                    ri = dgvNguoi.Rows.Add("tiền lấy", tienLayNguoi.ToString("N0"), "");
                    dgvNguoi.Rows[ri].DefaultCellStyle.BackColor = Color.White;
                    dgvNguoi.Rows[ri].Cells[1].Style.ForeColor = Color.Red;

                    // đơn trả (placeholder đỏ, tự điền)
                    ri = dgvNguoi.Rows.Add("đơn trả", "", "");
                    dgvNguoi.Rows[ri].DefaultCellStyle.ForeColor = Color.Red;

                    // đơn cũ ck (placeholder đỏ, tự điền)
                    ri = dgvNguoi.Rows.Add("đơn cũ ck", "", "");
                    dgvNguoi.Rows[ri].DefaultCellStyle.ForeColor = Color.Red;

                    // Dòng KẾT = TỔNG ĐƠN + tiền ship + tiền lấy (đơn trả/cũ ck để trống → không cộng)
                    decimal ketNguoi = tienThuNguoi + khoanShipNguoi + tienLayNguoi;
                    ri = dgvNguoi.Rows.Add("", ketNguoi.ToString("N0"), soDonNguoi.ToString("N0"));
                    dgvNguoi.Rows[ri].DefaultCellStyle.BackColor = AppConstants.COLOR_REPORT_KET;
                    dgvNguoi.Rows[ri].DefaultCellStyle.Font = new Font("Arial", 11, FontStyle.Bold);
                    dgvNguoi.Rows[ri].Height = AppConstants.ROW_HEIGHT_REPORT_KET;

                    pnlNguoi.Controls.Add(dgvNguoi);
                    pnlReports.Controls.Add(pnlNguoi);
                }
            }

            // Mở rộng pnlReports nếu nội dung vượt quá chiều rộng
            pnlReports.AutoScrollMinSize = new System.Drawing.Size(panelX, 0);
        }

        // ─── Invoice Button Panel ──────────────────────────────────────────────

        private void InitializeInvoiceButtonPanel()
        {
            Panel pnlButtons = tabInvoice.Controls["pnlInvoiceButtons"] as Panel;
            if (pnlButtons != null)
                return;

            pnlButtons = new Panel
            {
                Name = "pnlInvoiceButtons",
                BackColor = Color.FromArgb(40, 40, 40),
                Height = 40,
                Dock = DockStyle.Top,
            };
            tabInvoice.Controls.Add(pnlButtons);
            tabInvoice.Controls.SetChildIndex(pnlButtons, tabInvoice.Controls.Count - 1);

            Button MakeBtn(string text, int x) =>
                new Button
                {
                    Text = text,
                    BackColor = Color.FromArgb(40, 40, 40),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Arial", 9),
                    Size = new Size(75, 30),
                    Location = new Point(x, 5),
                };
            Button btnSave = MakeBtn("💾 Lưu", 10);
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += (s, e) => SaveDailyReportToExcel();
            Button btnUndo = MakeBtn("↶ Undo", 90);
            btnUndo.FlatAppearance.BorderSize = 0;
            btnUndo.Click += (s, e) => MessageBox.Show("↶ Undo thay đổi");
            Button btnClose = MakeBtn("✕ Đóng", 170);
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) =>
            {
                dgvInvoice.Rows.Clear();
                dgvInvoice.Columns.Clear();
                foreach (string name in new[] { "pnlDailyReport", "pnlInvoiceButtons" })
                {
                    var p = tabInvoice.Controls[name] as Panel;
                    if (p != null)
                    {
                        tabInvoice.Controls.Remove(p);
                        p.Dispose();
                    }
                }
            };

            pnlButtons.Controls.AddRange(new[] { btnSave, btnUndo, btnClose });
        }

        // ─── Save Daily Report → Excel ─────────────────────────────────────────

        private void SaveDailyReportToExcel()
        {
            try
            {
                if (string.IsNullOrEmpty(currentExcelFilePath))
                {
                    MessageBox.Show(
                        "Chưa mở file Excel. Vui lòng mở file Excel trước!",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }

                string sheetName =
                    tabExcelSheets.SelectedTab?.Text ?? DateTime.Now.ToString("dd-MM");
                DateTime sheetDate = DateTime.Now;
                DateTime.TryParseExact(
                    sheetName,
                    "dd-MM",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    out sheetDate
                );
                if (sheetDate.Year == 1)
                    sheetDate = sheetDate.AddYears(DateTime.Now.Year - 1);

                var service = new TextInputter.Services.ExcelInvoiceService(currentExcelFilePath);
                service.ApplyFormulasAndSummary(sheetName, sheetDate);

                MessageBox.Show(
                    $"✅ Đã ghi formula + bảng tổng kết vào:\n{System.IO.Path.GetFileName(currentExcelFilePath)}\nSheet: {sheetName}",
                    "✅ Lưu thành công",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                lblStatus.Text =
                    $"✅ Lưu formula → {System.IO.Path.GetFileName(currentExcelFilePath)} [{sheetName}]";
                lblStatus.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi lưu: {ex.Message}", "Lỗi");
                Debug.WriteLine($"Save error: {ex.Message}\n{ex.StackTrace}");
            }
        }

        // ─── Legacy handlers (buttons hidden in Designer, kept to avoid Designer wire errors) ──

        // NOTE: btnSaveInvoice, btnImportFromExcel, btnCalculateInvoice đều Visible=false trong Designer.
        // Flow chính dùng BtnCalculateExcelData_Click + SaveDailyReportToExcel thay thế.

        private void BtnSaveInvoice_Click(
            object sender,
            EventArgs e
        ) { /* hidden – dùng 💾 Lưu trong button panel */
        }

        private void BtnImportFromExcel_Click(
            object sender,
            EventArgs e
        ) { /* hidden – dùng BtnOpenExcel_Click + BtnCalculateExcelData_Click */
        }
    }
}

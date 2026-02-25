using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
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
            public string  Date         { get; set; }
            public decimal TongTienThu  { get; set; }   // Tổng tiền thu (cột H)
            public decimal TongTienShip { get; set; }   // Tổng tiền ship (cột I)
            public decimal KhoanTruShip { get; set; }   // -(TongShip - SoDon×5), số âm
            public decimal TongKetCuoi  { get; set; }   // TongTienThu + KhoanTruShip
            public decimal SoDon        { get; set; }
            // Các row âm (đơn trả, đơn cũ ck...) lấy từ Excel
            public List<(string Label, decimal Amount)> NegativeRows { get; set; } = new();
        }

        private DailyReportData currentDailyReport;

        // ─── Excel Viewer ──────────────────────────────────────────────────────

        private void BtnOpenExcel_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
                    openFileDialog.Title  = "Chọn file Excel";
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                        LoadExcelFile(openFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi:\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    if (sheetNames.Count == 0) { MessageBox.Show("⚠️ File Excel không có sheet nào"); return; }

                    tabExcelSheets.TabPages.Clear();

                    foreach (var sheetName in sheetNames)
                    {
                        TabPage tabPage = new TabPage(sheetName);
                        DataGridView dgv = new DataGridView
                        {
                            Dock                        = DockStyle.Fill,
                            AutoSizeColumnsMode         = DataGridViewAutoSizeColumnsMode.AllCells,
                            ReadOnly                    = false,
                            AllowUserToAddRows          = true,
                            AllowUserToDeleteRows       = true
                        };
                        tabPage.Controls.Add(dgv);
                        LoadSheetData(workbook, sheetName, dgv);
                        tabExcelSheets.TabPages.Add(tabPage);
                    }

                    tabMainControl.SelectedTab = tabExcelViewer;
                    lblStatus.Text      = $"✅ Excel: {System.IO.Path.GetFileName(filePath)} ({sheetNames.Count} sheets)";
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
                if (usedRange == null) return;

                int rowCount = usedRange.RowCount();
                int colCount = usedRange.ColumnCount();

                int headerRowIndex = 2;
                for (int row = 1; row <= Math.Min(5, rowCount); row++)
                {
                    string firstCell = worksheet.Cell(row, 1).GetString()?.Trim() ?? "";
                    if (firstCell == "SHOP" || firstCell.Contains("Tình trạng"))
                    { headerRowIndex = row; break; }
                }

                System.Data.DataTable dataTable = new System.Data.DataTable();
                for (int col = 1; col <= colCount; col++)
                    dataTable.Columns.Add(worksheet.Cell(headerRowIndex, col).GetString()?.Trim() ?? "");

                for (int row = 1; row <= rowCount; row++)
                {
                    if (row == headerRowIndex) continue;
                    var dataRow = dataTable.NewRow();
                    for (int col = 1; col <= colCount; col++)
                        dataRow[col - 1] = worksheet.Cell(row, col).GetString() ?? "";
                    dataTable.Rows.Add(dataRow);
                }

                dgv.DataSource = dataTable;
                dgv.AutoResizeColumns();
                if (dgv.Rows.Count > 0) dgv.Rows[0].Frozen = true;
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
                if (tabExcelSheets.TabPages.Count == 0) { MessageBox.Show("Chưa mở file Excel!"); return; }
                if (string.IsNullOrEmpty(currentExcelFilePath)) { MessageBox.Show("Không tìm thấy đường dẫn file Excel!", "Lỗi"); return; }

                using (var workbook = new XLWorkbook(currentExcelFilePath))
                {
                    foreach (TabPage tabPage in tabExcelSheets.TabPages)
                    {
                        var dgv = tabPage.Controls[0] as DataGridView;
                        if (dgv == null) continue;

                        var worksheet = workbook.Worksheet(tabPage.Text);
                        worksheet.Clear();

                        for (int col = 0; col < dgv.Columns.Count; col++)
                            worksheet.Cell(1, col + 1).Value = dgv.Columns[col].HeaderText;

                        for (int row = 0; row < dgv.Rows.Count; row++)
                            for (int col = 0; col < dgv.Columns.Count; col++)
                            {
                                var cellValue = dgv.Rows[row].Cells[col].Value;
                                if (cellValue != null) worksheet.Cell(row + 2, col + 1).Value = cellValue.ToString();
                            }
                    }
                    workbook.SaveAs(currentExcelFilePath);
                }

                MessageBox.Show("✅ Lưu file Excel thành công!", "Thành công");
                lblStatus.Text      = $"✅ Lưu Excel: {System.IO.Path.GetFileName(currentExcelFilePath)}";
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
                currentExcelFilePath    = "";
                lblStatus.Text          = "✅ Đã đóng file Excel";
                lblStatus.ForeColor     = Color.Green;
            }
            catch (Exception ex) { MessageBox.Show($"❌ Lỗi: {ex.Message}", "Lỗi"); }
        }

        // ─── Calculate (Excel → Daily Report) ─────────────────────────────────

        private void BtnCalculateExcelData_Click(object sender, EventArgs e)
        {
            try
            {
                if (tabExcelSheets.TabPages.Count == 0) return;

                var currentSheet = tabExcelSheets.SelectedTab;
                if (currentSheet == null || currentSheet.Controls.Count == 0) return;

                DataGridView sourceGridView = null;
                foreach (Control ctrl in currentSheet.Controls)
                    if (ctrl is DataGridView dgv) { sourceGridView = dgv; break; }

                if (sourceGridView == null || sourceGridView.Rows.Count == 0) return;

                // Column detection
                int colShop = -1, colTienThu = -1, colTienShip = -1, colTienHang = -1, colSoDon = -1, colGhiChu = -1, colNgayLay = -1;
                for (int col = 0; col < sourceGridView.Columns.Count; col++)
                {
                    string header = sourceGridView.Columns[col].HeaderText.ToLower();
                    if (header.Contains("shop"))       colShop     = col;
                    if (header.Contains("tiền thu"))   colTienThu  = col;
                    if (header.Contains("tiền ship"))  colTienShip = col;
                    if (header.Contains("tiền hàng"))  colTienHang = col;
                    if (header.Contains("số đơn"))     colSoDon    = col;
                    if (header.Contains("ghi chú"))    colGhiChu   = col;
                    if (header.Contains("ngày lấy"))   colNgayLay  = col;
                }

                Debug.WriteLine($"Cols — Shop:{colShop} TienThu:{colTienThu} TienShip:{colTienShip} TienHang:{colTienHang} SoDon:{colSoDon}");

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
                    if (sourceGridView.Columns[c].HeaderText.ToLower().Contains("mã")) { colMa = c; break; }

                // ── BƯỚC 1: Tìm SUM row trong Excel ────────────────────────────────
                decimal totalTienThu = 0, totalTienShip = 0, totalSoDon = 0;
                bool    foundSumRow  = false;
                int     sumRowIndex  = -1;

                for (int i = 0; i < sourceGridView.Rows.Count; i++)
                {
                    var row = sourceGridView.Rows[i];
                    if (row.IsNewRow) continue;
                    string shopVal = colShop >= 0 ? row.Cells[colShop].Value?.ToString() ?? "" : "";
                    if (!string.IsNullOrWhiteSpace(shopVal)) continue;

                    int checkCol = colTienThu >= 0 ? colTienThu : colTienHang;
                    if (checkCol < 0 || checkCol >= row.Cells.Count) continue;
                    if (!decimal.TryParse(row.Cells[checkCol].Value?.ToString() ?? "", out decimal chkVal) || chkVal <= 0) continue;

                    sumRowIndex = i;
                    foundSumRow = true;
                    if (colTienThu  >= 0) decimal.TryParse(row.Cells[colTienThu].Value?.ToString(),  out totalTienThu);
                    if (colTienShip >= 0) decimal.TryParse(row.Cells[colTienShip].Value?.ToString(), out totalTienShip);
                    if (colSoDon    >= 0) decimal.TryParse(row.Cells[colSoDon].Value?.ToString(),    out totalSoDon);
                    // Fallback: cột Column1 chứa SỐ ĐƠN khi header không detect được
                    if (totalSoDon == 0 && row.Cells.Count > AppConstants.COL_SODON_FALLBACK_IDX)
                        decimal.TryParse(row.Cells[AppConstants.COL_SODON_FALLBACK_IDX].Value?.ToString(), out totalSoDon);
                    // Log toàn bộ cells của SUM row để debug
                    var sbDebug = new System.Text.StringBuilder();
                    for (int dc = 0; dc < row.Cells.Count; dc++)
                        sbDebug.Append($"[{dc}]={row.Cells[dc].Value} ");
                    Debug.WriteLine($"SUM row idx={i}: {sbDebug}");
                    Debug.WriteLine($"SUM row idx={i}: TienThu={totalTienThu}, Ship={totalTienShip}, SoDon={totalSoDon}");
                    break;
                }

                // Nếu không có SUM row → tự cộng từng row DATA
                if (!foundSumRow)
                {
                    foreach (DataGridViewRow row in sourceGridView.Rows)
                    {
                        if (row.IsNewRow) continue;
                        string sv = colShop >= 0 ? row.Cells[colShop].Value?.ToString() ?? "" : "";
                        if (string.IsNullOrWhiteSpace(sv)) continue;
                        if (colMa >= 0 && colMa < row.Cells.Count && string.IsNullOrWhiteSpace(row.Cells[colMa].Value?.ToString() ?? "")) continue;
                        if (colTienThu  >= 0) { if (decimal.TryParse(row.Cells[colTienThu].Value?.ToString(),  out decimal t)) totalTienThu  += t; }
                        if (colTienShip >= 0) { if (decimal.TryParse(row.Cells[colTienShip].Value?.ToString(), out decimal s)) totalTienShip += s; }
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
                        if (row.IsNewRow) continue;
                        if (colTienHangCheck >= row.Cells.Count) continue;
                        if (!decimal.TryParse(row.Cells[colTienHangCheck].Value?.ToString() ?? "", out decimal jVal) || jVal >= 0) continue;

                        // Loại bỏ nếu có MÃ HĐ (đơn thật bị âm, không phải khoản trừ)
                        if (colMa >= 0 && colMa < row.Cells.Count && !string.IsNullOrWhiteSpace(row.Cells[colMa].Value?.ToString())) continue;
                        // Loại bỏ nếu có SHOP (đơn thật bị âm, không phải khoản trừ)
                        if (colShop >= 0 && colShop < row.Cells.Count && !string.IsNullOrWhiteSpace(row.Cells[colShop].Value?.ToString())) continue;

                        negativeRows.Add(row);
                    }
                }

                // Tính tổng số âm ở TIỀN HÀNG
                decimal totalNegHang = 0;
                foreach (var nr in negativeRows)
                    if (decimal.TryParse(nr.Cells[colTienHangCheck].Value?.ToString() ?? "", out decimal nv)) totalNegHang += nv;

                decimal tongHangDuong = totalTienThu - totalTienShip;        // SUM row TIỀN HÀNG
                decimal tongKetCuoi   = tongHangDuong + totalNegHang;        // cộng luôn số âm
                decimal phiShipThucTe = totalSoDon * AppConstants.PHI_SHIP_MOI_DON;
                decimal khoanTruShip  = -(totalTienShip - phiShipThucTe);

                Debug.WriteLine($"FINAL: SumRow={foundSumRow}, Thu={totalTienThu}, Ship={totalTienShip}, HangDuong={tongHangDuong}, NegHang={totalNegHang}, KetCuoi={tongKetCuoi}");

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
                            dgvInvoice.Rows[idx].Cells[ci].Style.Font = new Font(dgvInvoice.Font, FontStyle.Italic);
                }

                // 1. Data rows (có SHOP và có MÃ)
                for (int i = 0; i < (sumRowIndex >= 0 ? sumRowIndex : sourceGridView.Rows.Count); i++)
                {
                    var sr = sourceGridView.Rows[i];
                    if (sr.IsNewRow) continue;
                    string sv = colShop >= 0 ? sr.Cells[colShop].Value?.ToString() ?? "" : "";
                    if (string.IsNullOrWhiteSpace(sv)) continue;
                    if (colMa >= 0 && colMa < sr.Cells.Count && string.IsNullOrWhiteSpace(sr.Cells[colMa].Value?.ToString() ?? "")) continue;
                    AddRow(sr, null, false);
                }

                // 2. SUM row — màu vàng
                {
                    var sumRow = new DataGridViewRow();
                    sumRow.CreateCells(dgvInvoice);
                    if (sumRow.Cells.Count > 0) sumRow.Cells[0].Value = "▶ TỔNG";
                    if (colTienThu  >= 0 && colTienThu  < sumRow.Cells.Count) sumRow.Cells[colTienThu].Value  = totalTienThu.ToString();
                    if (colTienShip >= 0 && colTienShip < sumRow.Cells.Count) sumRow.Cells[colTienShip].Value = totalTienShip.ToString();
                    if (colTienHang >= 0 && colTienHang < sumRow.Cells.Count) sumRow.Cells[colTienHang].Value = tongHangDuong.ToString();
                    if (colSoDon    >= 0 && colSoDon    < sumRow.Cells.Count) sumRow.Cells[colSoDon].Value    = totalSoDon.ToString();
                    // Không ghi fallback vào cells[16] vì sẽ đè vào cột FAIL
                    dgvInvoice.Rows.Add(sumRow);
                    int si = dgvInvoice.Rows.Count - 1;
                    for (int ci = 0; ci < dgvInvoice.Columns.Count; ci++)
                    {
                        dgvInvoice.Rows[si].Cells[ci].Style.BackColor = AppConstants.COLOR_ROW_TONG;
                        dgvInvoice.Rows[si].Cells[ci].Style.ForeColor = Color.Black;
                        dgvInvoice.Rows[si].Cells[ci].Style.Font      = new Font(dgvInvoice.Font, FontStyle.Bold);
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
                    if (ketRow.Cells.Count > 0) ketRow.Cells[0].Value = "▶ KẾT";
                    if (colTienHang >= 0 && colTienHang < ketRow.Cells.Count) ketRow.Cells[colTienHang].Value = tongKetCuoi.ToString();
                    if (colSoDon >= 0 && colSoDon < ketRow.Cells.Count) ketRow.Cells[colSoDon].Value = totalSoDon.ToString();
                    // Fallback cột fallback index nếu không detect colSoDon
                    if (colSoDon < 0 && ketRow.Cells.Count > AppConstants.COL_SODON_FALLBACK_IDX)
                        ketRow.Cells[AppConstants.COL_SODON_FALLBACK_IDX].Value = totalSoDon.ToString();
                    dgvInvoice.Rows.Add(ketRow);
                    int ki = dgvInvoice.Rows.Count - 1;
                    for (int ci = 0; ci < dgvInvoice.Columns.Count; ci++)
                    {
                        dgvInvoice.Rows[ki].Cells[ci].Style.BackColor = AppConstants.COLOR_ROW_KET;
                        dgvInvoice.Rows[ki].Cells[ci].Style.ForeColor = Color.Black;
                        dgvInvoice.Rows[ki].Cells[ci].Style.Font = new Font(dgvInvoice.Font, FontStyle.Bold);
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
                    Date         = reportDate,
                    TongTienThu  = totalTienThu,
                    TongTienShip = totalTienShip,
                    KhoanTruShip = khoanTruShip,
                    TongKetCuoi  = tongKetCuoi,
                    SoDon        = totalSoDon,
                    NegativeRows = negativeRows.Select(nr =>
                    {
                        // Tìm label: quét tất cả cells, lấy ô có text (không phải số, không rỗng)
                        string lbl = "";
                        for (int ci = 0; ci < nr.Cells.Count; ci++)
                        {
                            string v = nr.Cells[ci].Value?.ToString()?.Trim() ?? "";
                            if (string.IsNullOrEmpty(v)) continue;
                            if (decimal.TryParse(v, out _)) continue; // bỏ qua ô số
                            lbl = v;
                            break;
                        }
                        if (string.IsNullOrEmpty(lbl)) lbl = "đơn âm";
                        decimal.TryParse(nr.Cells[colTienHangCheck].Value?.ToString() ?? "", out decimal amt);
                        return (lbl, amt);
                    }).ToList()
                };

                lblInvoiceTotal.Text = $"TỔNG THU: {totalTienThu:N0} đ | SHIP: {totalTienShip:N0} đ | SỐ ĐƠN: {totalSoDon:N0} | KẾT: {tongKetCuoi:N0} đ";

                DisplayDailyReport();
                InitializeInvoiceButtonPanel();
                tabMainControl.SelectedIndex = 2;
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
                dgvInvoice.Columns.Add("Tên",    "Tên");
                dgvInvoice.Columns.Add("Tiền",   "Tiền");
                dgvInvoice.Columns.Add("Số đơn", "Số đơn");
            }
            dgvInvoice.Rows.Add("", "0", "0");
        }

        private void BtnCalculateInvoice_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvInvoice.Rows.Count == 0) { MessageBox.Show("Chưa có dữ liệu để tính!"); return; }

                decimal totalTien = 0, totalSoDon = 0;
                for (int i = 0; i < dgvInvoice.Rows.Count; i++)
                {
                    if (decimal.TryParse(dgvInvoice.Rows[i].Cells[1].Value?.ToString() ?? "0", out decimal tienHang))
                        totalTien += tienHang;
                    if (decimal.TryParse(dgvInvoice.Rows[i].Cells.Count > 8
                            ? dgvInvoice.Rows[i].Cells[8].Value?.ToString() ?? "0" : "0", out decimal sodon))
                        totalSoDon += sodon;
                }

                lblInvoiceTotal.Text = $"TỔNG CỘNG: {totalTien:N0} đ | SỐ ĐƠN: {totalSoDon:N0}";

                currentDailyReport = new DailyReportData
                {
                    Date         = DateTime.Now.ToString("dd.MM.yyyy"),
                    TongTienThu  = totalTien,
                    TongTienShip = 0,
                    KhoanTruShip = 0,
                    TongKetCuoi  = totalTien,
                    SoDon        = totalSoDon
                };

                InitializeInvoiceButtonPanel();
                DisplayDailyReport();
            }
            catch (Exception ex) { Debug.WriteLine($"❌ Lỗi: {ex.Message}"); }
        }

        // ─── Daily Report Display ──────────────────────────────────────────────

        private void DisplayDailyReport()
        {
            if (currentDailyReport == null) return;

            Panel pnlTop    = tabInvoice.Controls["pnlInvoiceTop"]          as Panel;
            Panel pnlBottom = tabInvoice.Controls["pnlDailyReportBottom"]   as Panel;

            if (pnlTop == null)
            {
                tabInvoice.Controls.Clear();

                pnlTop = new Panel { Name = "pnlInvoiceTop", Dock = DockStyle.Fill, BackColor = Color.White };
                pnlTop.Controls.Add(dgvInvoice);
                pnlTop.Controls.Add(lblInvoiceTotal);
                tabInvoice.Controls.Add(pnlTop);

                pnlBottom = new Panel
                {
                    Name        = "pnlDailyReportBottom",
                    Dock        = DockStyle.Bottom,
                    BackColor   = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    Height      = AppConstants.DAILY_REPORT_PANEL_HEIGHT
                };
                tabInvoice.Controls.Add(pnlBottom);
            }

            pnlBottom.Controls.Clear();

            var r = currentDailyReport;
            string soDonStr   = r.SoDon.ToString("N0");
            string thuStr     = r.TongTienThu.ToString("N0");
            string shipTruStr = r.KhoanTruShip.ToString("N0");
            string ketStr     = r.TongKetCuoi.ToString("N0");

            Debug.WriteLine($"DisplayDailyReport: TongThu={r.TongTienThu}, TongShip={r.TongTienShip}, KhoanTru={r.KhoanTruShip}, TongKet={r.TongKetCuoi}, SoDon={r.SoDon}");

            var dgvReport = new DataGridView
            {
                Dock                  = DockStyle.Fill,
                BackgroundColor       = Color.White,
                AllowUserToAddRows    = false,
                AllowUserToDeleteRows = false,
                ReadOnly              = true,
                ColumnHeadersVisible  = false,
                RowHeadersVisible     = false,
                ScrollBars            = ScrollBars.Both,
                DefaultCellStyle      = { Font = new Font("Arial", 10), Alignment = DataGridViewContentAlignment.MiddleLeft }
            };

            dgvReport.Columns.Add("TenMuc", "");
            dgvReport.Columns.Add("Tien",   "");
            dgvReport.Columns.Add("SoDon",  "");
            dgvReport.Columns[0].Width = 220;
            dgvReport.Columns[1].Width = 110;
            dgvReport.Columns[2].Width = 90;
            dgvReport.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvReport.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            int ri;

            ri = dgvReport.Rows.Add("", "Tiền Thu", "Số đơn");
            dgvReport.Rows[ri].DefaultCellStyle.BackColor = Color.LightSteelBlue;
            dgvReport.Rows[ri].DefaultCellStyle.Font      = new Font("Arial", 10, FontStyle.Bold);

            ri = dgvReport.Rows.Add("TỔNG ĐƠN", thuStr, soDonStr);
            dgvReport.Rows[ri].DefaultCellStyle.BackColor = Color.White;

            ri = dgvReport.Rows.Add("tiền ship", shipTruStr, "");
            dgvReport.Rows[ri].DefaultCellStyle.BackColor = Color.White;
            dgvReport.Rows[ri].Cells[1].Style.ForeColor   = r.KhoanTruShip < 0 ? Color.Red : Color.Black;

            dgvReport.Rows.Add("tiền lấy",  "", "");

            // Render các row âm động từ Excel (đơn trả, đơn cũ ck...)
            foreach (var (label, amount) in r.NegativeRows)
            {
                ri = dgvReport.Rows.Add(label, amount.ToString("N0"), "");
                dgvReport.Rows[ri].DefaultCellStyle.ForeColor = Color.Red;
            }

            dgvReport.Rows.Add("", "", "");

            ri = dgvReport.Rows.Add("", ketStr, soDonStr);
            dgvReport.Rows[ri].DefaultCellStyle.BackColor = AppConstants.COLOR_REPORT_KET;
            dgvReport.Rows[ri].DefaultCellStyle.Font      = new Font("Arial", 11, FontStyle.Bold);
            dgvReport.Rows[ri].Height = AppConstants.ROW_HEIGHT_REPORT_KET;

            pnlBottom.Controls.Add(dgvReport);
        }

        // ─── Invoice Button Panel ──────────────────────────────────────────────

        private void InitializeInvoiceButtonPanel()
        {
            Panel pnlButtons = tabInvoice.Controls["pnlInvoiceButtons"] as Panel;
            if (pnlButtons != null) return;

            pnlButtons = new Panel
            {
                Name      = "pnlInvoiceButtons",
                BackColor = Color.FromArgb(40, 40, 40),
                Height    = 40,
                Dock      = DockStyle.Top
            };
            tabInvoice.Controls.Add(pnlButtons);
            tabInvoice.Controls.SetChildIndex(pnlButtons, tabInvoice.Controls.Count - 1);

            Button MakeBtn(string text, int x) => new Button
            {
                Text        = text,
                BackColor   = Color.FromArgb(40, 40, 40),
                ForeColor   = Color.White,
                FlatStyle   = FlatStyle.Flat,
                Font        = new Font("Arial", 9),
                Size        = new Size(75, 30),
                Location    = new Point(x, 5)
            };
            Button btnSave    = MakeBtn("💾 Lưu",   10);  btnSave.FlatAppearance.BorderSize  = 0; btnSave.Click  += (s, e) => SaveDailyReportToExcel();
            Button btnUndo    = MakeBtn("↶ Undo",   90);  btnUndo.FlatAppearance.BorderSize  = 0; btnUndo.Click  += (s, e) => MessageBox.Show("↶ Undo thay đổi");
            Button btnClose   = MakeBtn("✕ Đóng",  170);  btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) =>
            {
                dgvInvoice.Rows.Clear();
                dgvInvoice.Columns.Clear();
                foreach (string name in new[] { "pnlDailyReport", "pnlInvoiceButtons" })
                {
                    var p = tabInvoice.Controls[name] as Panel;
                    if (p != null) { tabInvoice.Controls.Remove(p); p.Dispose(); }
                }
            };

            pnlButtons.Controls.AddRange(new[] { btnSave, btnUndo, btnClose });
        }

        // ─── Save Daily Report → Excel ─────────────────────────────────────────

        private void SaveDailyReportToExcel()
        {
            try
            {
                if (dgvInvoice.Rows.Count == 0) { MessageBox.Show("Không có dữ liệu để lưu!"); return; }

                // Chọn file lưu qua dialog (không hardcode)
                string defaultName = (currentDailyReport?.Date is string d2 && !string.IsNullOrEmpty(d2))
                    ? $"DailyReport_{d2.Replace("/", "-").Replace(".", "-")}.xlsx"
                    : AppConstants.DAILY_REPORT_FILENAME;
                string excelPath;
                using (var sfd = new SaveFileDialog
                {
                    Title            = "Lưu báo cáo hàng ngày",
                    Filter           = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                    DefaultExt       = "xlsx",
                    FileName         = defaultName,
                    InitialDirectory = System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(currentExcelFilePath ?? ""))
                                       ? System.IO.Path.GetDirectoryName(currentExcelFilePath)
                                       : AppDomain.CurrentDomain.BaseDirectory
                })
                {
                    if (sfd.ShowDialog() != DialogResult.OK) return;
                    excelPath = sfd.FileName;
                }
                // Sheet name = ngày lấy từ data; fallback = hôm nay
                string sheetName = (currentDailyReport?.Date is string d && !string.IsNullOrEmpty(d))
                    ? d
                    : DateTime.Now.ToString(AppConstants.DATE_FORMAT_SHEET);

                XLWorkbook workbook;
                if (System.IO.File.Exists(excelPath))
                {
                    workbook = new XLWorkbook(excelPath);
                    var existingSheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
                    if (existingSheet != null) workbook.Worksheets.Delete(sheetName);
                }
                else workbook = new XLWorkbook();

                using (workbook)
                {
                    var worksheet  = workbook.Worksheets.Add(sheetName);
                    int currentRow = 1;

                    // Phần 1: Invoice data
                    for (int col = 0; col < dgvInvoice.Columns.Count; col++)
                    {
                        worksheet.Cell(currentRow, col + 1).Value = dgvInvoice.Columns[col].HeaderText;
                        worksheet.Cell(currentRow, col + 1).Style.Font.Bold = true;
                        worksheet.Cell(currentRow, col + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                    }
                    currentRow++;

                    for (int row = 0; row < dgvInvoice.Rows.Count; row++)
                    {
                        for (int col = 0; col < dgvInvoice.Columns.Count; col++)
                        {
                            var cellValue = dgvInvoice.Rows[row].Cells[col].Value;
                            worksheet.Cell(currentRow, col + 1).Value = cellValue?.ToString() ?? "";
                            if (row == dgvInvoice.Rows.Count - 1)
                            {
                                worksheet.Cell(currentRow, col + 1).Style.Font.Bold = true;
                                worksheet.Cell(currentRow, col + 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
                            }
                        }
                        currentRow++;
                    }

                    currentRow += 2;

                    // Phần 2: Daily Report — ghi thẳng từ currentDailyReport (không đọc UI)
                    if (currentDailyReport != null)
                    {
                        var r = currentDailyReport;

                        // Tiêu đề phần 2
                        worksheet.Cell(currentRow, 1).Value = "BÁO CÁO HÀNG NGÀY";
                        worksheet.Cell(currentRow, 1).Style.Font.Bold     = true;
                        worksheet.Cell(currentRow, 1).Style.Font.FontSize = 12;
                        currentRow++;

                        // Header row
                        worksheet.Cell(currentRow, 1).Value = "";
                        worksheet.Cell(currentRow, 2).Value = "Tiền Thu";
                        worksheet.Cell(currentRow, 3).Value = "Số đơn";
                        for (int c = 1; c <= 3; c++)
                        {
                            worksheet.Cell(currentRow, c).Style.Font.Bold = true;
                            worksheet.Cell(currentRow, c).Style.Fill.BackgroundColor = XLColor.LightSteelBlue;
                        }
                        currentRow++;

                        // TỔNG ĐƠN
                        worksheet.Cell(currentRow, 1).Value = "TỔNG ĐƠN";
                        worksheet.Cell(currentRow, 2).Value = r.TongTienThu.ToString("N0");
                        worksheet.Cell(currentRow, 3).Value = r.SoDon.ToString("N0");
                        currentRow++;

                        // tiền ship
                        worksheet.Cell(currentRow, 1).Value = "tiền ship";
                        worksheet.Cell(currentRow, 2).Value = r.KhoanTruShip.ToString("N0");
                        currentRow++;

                        // tiền lấy
                        worksheet.Cell(currentRow, 1).Value = "tiền lấy";
                        currentRow++;

                        // Các row âm (đơn trả, đơn cũ ck...)
                        foreach (var (label, amount) in r.NegativeRows)
                        {
                            worksheet.Cell(currentRow, 1).Value = label;
                            worksheet.Cell(currentRow, 2).Value = amount.ToString("N0");
                            worksheet.Cell(currentRow, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(currentRow, 2).Style.Font.FontColor = XLColor.Red;
                            currentRow++;
                        }

                        // Dòng trống
                        currentRow++;

                        // Dòng KẾT (tổng kết)
                        worksheet.Cell(currentRow, 2).Value = r.TongKetCuoi.ToString("N0");
                        worksheet.Cell(currentRow, 3).Value = r.SoDon.ToString("N0");
                        for (int c = 1; c <= 3; c++)
                        {
                            worksheet.Cell(currentRow, c).Style.Font.Bold = true;
                            worksheet.Cell(currentRow, c).Style.Fill.BackgroundColor = XLColor.Orange;
                            worksheet.Cell(currentRow, c).Style.Font.FontSize = 11;
                        }
                        currentRow++;
                    }

                    worksheet.Columns().AdjustToContents();
                    workbook.SaveAs(excelPath);
                }

                MessageBox.Show($"✅ Lưu thành công vào:\n{excelPath}\n\nSheet: {sheetName}\n\n✓ Phần 1 (Invoice)\n✓ Phần 2 (Daily Report)", "Thành công");
                lblStatus.Text      = $"✅ Lưu Daily Report: {sheetName}";
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

        private void BtnSaveInvoice_Click(object sender, EventArgs e) { /* hidden – dùng 💾 Lưu trong button panel */ }

        private void BtnImportFromExcel_Click(object sender, EventArgs e) { /* hidden – dùng BtnOpenExcel_Click + BtnCalculateExcelData_Click */ }
    }
}

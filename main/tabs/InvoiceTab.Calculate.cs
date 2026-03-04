using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TextInputter
{
    // ─── Calculate: Excel → Daily Report ────────────────────────────────────────
    public partial class MainForm
    {
        private async void BtnCalculateExcelData_Click(object sender, EventArgs e)
        {
            // Disable nút + hiện loading overlay ngay lập tức để user biết app đang xử lý
            btnCalculateExcelData.Enabled = false;
            Panel overlay = ShowLoadingOverlay("⏳ Đang tính tiền...");
            try
            {
                // Nhường control loop 1 frame để overlay render trước khi bắt đầu tính
                await Task.Delay(30);

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
                    // Không đọc totalSoDon từ SUM row vì cột "số đơn" thường không có trong source Excel.
                    // Sẽ đếm chính xác từ data rows bên dưới.
                    Debug.WriteLine(
                        $"SUM row idx={i}: TienThu={totalTienThu}, Ship={totalTienShip}"
                    );
                    break;
                }

                // Luôn đếm số đơn từ data rows thực tế (không dùng SUM row) để đảm bảo chính xác
                int endCountIdx = sumRowIndex >= 0 ? sumRowIndex : sourceGridView.Rows.Count;
                for (int i = 0; i < endCountIdx; i++)
                {
                    var row = sourceGridView.Rows[i];
                    if (row.IsNewRow)
                        continue;
                    string sv = colShop >= 0 ? row.Cells[colShop].Value?.ToString() ?? "" : "";
                    if (string.IsNullOrWhiteSpace(sv))
                        continue;
                    if (IsDateLabelRow(row, colShop, colMa))
                        continue;
                    totalSoDon++;
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
                        if (IsDateLabelRow(row, colShop, colMa))
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
                        if (IsDateLabelRow(row, colShop, colMa))
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
                    if (IsDateLabelRow(sr, colShop, colMa))
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
                MessageBox.Show(
                    $"❌ Lỗi khi tính tiền:\n{ex.Message}",
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                HideLoadingOverlay(overlay);
                btnCalculateExcelData.Enabled = true;
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
    }
}

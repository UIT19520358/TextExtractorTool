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
                    colNguoiDi = -1,
                    colTenKH = -1,
                    colDiaChi = -1,
                    colQuan = -1,
                    colUngTien = -1,
                    colHangTon = -1,
                    colFail = -1;
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
                    if (header.Contains("tên kh"))
                        colTenKH = col;
                    if (header.Contains("địa chỉ") || header.Contains("dia chi"))
                        colDiaChi = col;
                    if (header.Contains("quận") || header.Contains("quan"))
                        colQuan = col;
                    if (header.Contains("ứng tiền") || header.Contains("ung tien"))
                        colUngTien = col;
                    if (header.Contains("hàng tồn") || header.Contains("hang ton"))
                        colHangTon = col;
                    if (header.Contains("fail"))
                        colFail = col;
                }

                Debug.WriteLine(
                    $"Cols — Shop:{colShop} TienThu:{colTienThu} TienShip:{colTienShip} TienHang:{colTienHang} SoDon:{colSoDon} TenKH:{colTenKH} DiaChi:{colDiaChi} Quan:{colQuan} Fail:{colFail}"
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

                // ── BƯỚC 1: Tìm SUM row trong Excel (dùng làm mốc cuối data) ────────
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
                    Debug.WriteLine(
                        $"SUM row idx={i}: detected (dùng làm mốc cuối data, ko lấy giá trị)"
                    );
                    break;
                }

                // LUÔN tính THU + SHIP từ từng row DATA (không dùng SUM row)
                // Đảm bảo UI giống exported Excel (cũng dùng SUBTOTAL per-row).
                // Track riêng hàng tồn (HÀNG TỒN=x) để loại khỏi LEFT summary.
                decimal totalTienThuHangTon = 0, totalTienShipHangTon = 0;
                int soDonHangTon = 0;
                // LEFT "Đơn trả": TIỀN HÀNG of FAIL=xx AND ỨNG TIỀN=x (matching Excel formula)
                decimal totalTienHangDonTra = 0;
                int soDonTraLeft = 0;

                int endDataIdx = sumRowIndex >= 0 ? sumRowIndex : sourceGridView.Rows.Count;
                for (int i = 0; i < endDataIdx; i++)
                {
                    var row = sourceGridView.Rows[i];
                    if (row.IsNewRow)
                        continue;
                    string sv = colShop >= 0 ? row.Cells[colShop].Value?.ToString() ?? "" : "";
                    if (string.IsNullOrWhiteSpace(sv))
                        continue;
                    if (IsDateLabelRow(row, colShop, colMa))
                        continue;

                    // Detect hàng tồn (carry-over ngày trước)
                    bool isHangTon = false;
                    if (colHangTon >= 0 && colHangTon < row.Cells.Count)
                    {
                        string tonVal = (row.Cells[colHangTon].Value?.ToString() ?? "").Trim().ToLower();
                        isHangTon = tonVal == "x";
                    }

                    // Detect đơn trả cho LEFT summary: FAIL=xx AND ỨNG TIỀN=x → sum TIỀN HÀNG
                    bool isTraLeft = false;
                    if (colFail >= 0 && colFail < row.Cells.Count && colUngTien >= 0 && colUngTien < row.Cells.Count)
                    {
                        string failVal = (row.Cells[colFail].Value?.ToString() ?? "").Trim().ToLower();
                        string ungVal = (row.Cells[colUngTien].Value?.ToString() ?? "").Trim().ToLower();
                        isTraLeft = failVal.Contains("xx") && ungVal == "x";
                    }
                    if (isTraLeft && colTienHang >= 0 && colTienHang < row.Cells.Count)
                    {
                        if (decimal.TryParse(row.Cells[colTienHang].Value?.ToString() ?? "", out decimal hangVal))
                            totalTienHangDonTra += hangVal;
                        soDonTraLeft++;
                    }

                    totalSoDon++;
                    if (isHangTon)
                        soDonHangTon++;

                    if (colTienThu >= 0 && colTienThu < row.Cells.Count)
                        if (
                            decimal.TryParse(
                                row.Cells[colTienThu].Value?.ToString() ?? "",
                                out decimal t
                            )
                        )
                        {
                            totalTienThu += t;
                            if (isHangTon)
                                totalTienThuHangTon += t;
                        }

                    if (colTienShip >= 0 && colTienShip < row.Cells.Count)
                        if (
                            decimal.TryParse(
                                row.Cells[colTienShip].Value?.ToString() ?? "",
                                out decimal s
                            )
                        )
                        {
                            totalTienShip += s;
                            if (isHangTon)
                                totalTienShipHangTon += s;
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

                // ── Tổng hợp theo NGƯỜI ĐI (+ detect đơn gộp, đơn trả) ────────
                var detailByNguoiDi = new Dictionary<string, NguoiDiDetail>(StringComparer.OrdinalIgnoreCase);
                int totalDonGop = 0, totalDonTra = 0;

                // Struct tạm collect row data per người đi (dùng cho gộp detection + đơn trả details)
                var rowsPerNguoi = new Dictionary<string, List<(string TenKH, string DiaChi, string Quan, string Ma, decimal TienThu, decimal ShipFee, bool IsTra)>>(StringComparer.OrdinalIgnoreCase);

                {
                    int endIdx = sumRowIndex >= 0 ? sumRowIndex : sourceGridView.Rows.Count;
                    for (int i = 0; i < endIdx; i++)
                    {
                        var row = sourceGridView.Rows[i];
                        if (row.IsNewRow)
                            continue;
                        string sv = colShop >= 0 ? row.Cells[colShop].Value?.ToString() ?? "" : "";
                        if (string.IsNullOrWhiteSpace(sv))
                            continue;
                        if (IsDateLabelRow(row, colShop, colMa))
                            continue;

                        string nguoiRow = colNguoiDi >= 0 && colNguoiDi < row.Cells.Count
                            ? (row.Cells[colNguoiDi].Value?.ToString() ?? "").Trim() : "";
                        if (string.IsNullOrEmpty(nguoiRow))
                            nguoiRow = "(không rõ)";

                        decimal tienThuRow = 0;
                        if (colTienThu >= 0 && colTienThu < row.Cells.Count)
                            decimal.TryParse(row.Cells[colTienThu].Value?.ToString() ?? "", out tienThuRow);

                        decimal tienShipRow = 0;
                        if (colTienShip >= 0 && colTienShip < row.Cells.Count)
                            decimal.TryParse(row.Cells[colTienShip].Value?.ToString() ?? "", out tienShipRow);

                        string tenKH = colTenKH >= 0 && colTenKH < row.Cells.Count
                            ? (row.Cells[colTenKH].Value?.ToString() ?? "").Trim() : "";
                        string diaChi = colDiaChi >= 0 && colDiaChi < row.Cells.Count
                            ? (row.Cells[colDiaChi].Value?.ToString() ?? "").Trim() : "";
                        string quan = colQuan >= 0 && colQuan < row.Cells.Count
                            ? (row.Cells[colQuan].Value?.ToString() ?? "").Trim() : "";
                        string ma = colMa >= 0 && colMa < row.Cells.Count
                            ? (row.Cells[colMa].Value?.ToString() ?? "").Trim() : "";

                        // Detect đơn trả: FAIL = "xx"
                        bool isTra = false;
                        if (colFail >= 0 && colFail < row.Cells.Count)
                        {
                            string failVal = (row.Cells[colFail].Value?.ToString() ?? "").Trim().ToLower();
                            isTra = failVal.Contains("xx");
                        }

                        // Tích lũy per-person
                        if (!detailByNguoiDi.ContainsKey(nguoiRow))
                            detailByNguoiDi[nguoiRow] = new NguoiDiDetail();
                        var d = detailByNguoiDi[nguoiRow];
                        d.TienThu += tienThuRow;
                        d.TienShip += tienShipRow;
                        d.SoDon++;
                        d.IsAnTam = nguoiRow.StartsWith(AppConstants.NGUOI_DI_DEFAULT  + DateTime.Now.ToString(" dd-MM"), StringComparison.OrdinalIgnoreCase);
                        if (isTra)
                            d.SoDonTra++;

                        // Collect row data cho gộp detection + đơn trả details
                        if (!rowsPerNguoi.ContainsKey(nguoiRow))
                            rowsPerNguoi[nguoiRow] = new List<(string, string, string, string, decimal, decimal, bool)>();
                        rowsPerNguoi[nguoiRow].Add((tenKH, diaChi, quan, ma, tienThuRow, tienShipRow, isTra));
                    }
                }

                // Detect đơn gộp + calculate deductions per người đi
                foreach (var kvp in detailByNguoiDi)
                {
                    string nguoi = kvp.Key;
                    var d = kvp.Value;
                    if (!rowsPerNguoi.ContainsKey(nguoi))
                        continue;
                    var rows = rowsPerNguoi[nguoi];

                    // Đơn gộp: cùng TÊN KH + ĐỊA CHỈ → giao 1 lần, nhóm >1 = gộp
                    var groups = rows
                        .Where(r => !string.IsNullOrEmpty(r.TenKH) && !string.IsNullOrEmpty(r.DiaChi))
                        .GroupBy(r => (r.TenKH.ToLower(), r.DiaChi.ToLower()));
                    foreach (var g in groups)
                        if (g.Count() > 1)
                            d.SoDonGop += g.Count() - 1;
                    totalDonGop += d.SoDonGop;
                    totalDonTra += d.SoDonTra;

                    // Tính deductions (skip An Tâm — không trừ ship/lấy/trả)
                    if (!d.IsAnTam)
                    {
                        // Tiền ship trừ: -(TongShip - SoDonGiao × 5k)
                        d.TienShipTru = -(d.TienShip - d.SoDonGiao * AppConstants.PHI_SHIP_MOI_DON);

                        // Tiền lấy: -((SoDon - SoDonTra - SoDonGop) × 2k)
                        decimal donLayThucTe = d.SoDon - d.SoDonTra - d.SoDonGop;
                        if (donLayThucTe < 0)
                            donLayThucTe = 0;
                        d.TienLay = -(donLayThucTe * AppConstants.PHI_LAY_HANG_MOI_DON);

                        // Đơn trả: -(tienThu - shipFee + 5k) per return
                        foreach (var r in rows.Where(r => r.IsTra))
                        {
                            decimal shipFeeLookup = LookupShipFeeByQuan(r.Quan);
                            decimal deduction = -(r.TienThu - shipFeeLookup + AppConstants.PHI_CONG_DON_TRA);
                            d.TienDonTra += deduction;
                            d.DonTraDetails.Add((r.Ma, r.TienThu, shipFeeLookup, deduction));
                        }
                    }
                }

                Debug.WriteLine(
                    $"FINAL: SumRow={foundSumRow}, Thu={totalTienThu}, Ship={totalTienShip}, HangDuong={tongHangDuong}, NegHang={totalNegHang}, KetCuoi={tongKetCuoi}, DonGop={totalDonGop}, DonTra={totalDonTra}"
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
                    TotalDonGop = totalDonGop,
                    TotalDonTra = totalDonTra,
                    TongTienHangDonTra = totalTienHangDonTra,
                    SoDonTraLeft = soDonTraLeft,
                    TongTienThuHangTon = totalTienThuHangTon,
                    TongTienShipHangTon = totalTienShipHangTon,
                    SoDonHangTon = soDonHangTon,
                    DetailByNguoiDi = detailByNguoiDi,
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

                // Giới hạn chiều rộng cột (ĐỊA CHỈ quá dài đẩy mất cột đầu)
                foreach (DataGridViewColumn col in dgvInvoice.Columns)
                {
                    if (col.Width > 300)
                        col.Width = 300;
                }
                // Scroll về cột đầu tiên để ko bị mất mấy cột bên trái
                if (dgvInvoice.Columns.Count > 0)
                    dgvInvoice.FirstDisplayedScrollingColumnIndex = 0;
                // Scroll về hàng đầu tiên để ko bị che data phía trên
                if (dgvInvoice.Rows.Count > 0)
                    dgvInvoice.FirstDisplayedScrollingRowIndex = 0;

                DisplayDailyReport();
                InitializeInvoiceButtonPanel();
                tabMainControl.SelectedIndex = 2;
                tabInvoice.PerformLayout();

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

        // ─── Ship fee lookup helper ────────────────────────────────────────────

        /// <summary>
        /// Tra cứu phí ship theo QUẬN từ AppConstants.SHIPPING_FEES_BY_QUAN.
        /// Tự bỏ dấu tiếng Việt + normalize trước khi tra.
        /// Trả về 0 nếu không tìm thấy.
        /// </summary>
        private static decimal LookupShipFeeByQuan(string quan)
        {
            if (string.IsNullOrWhiteSpace(quan))
                return 0m;

            // 1. Thử exact match (case-insensitive đã có trong dictionary)
            if (AppConstants.SHIPPING_FEES_BY_QUAN.TryGetValue(quan.Trim(), out decimal fee1))
                return fee1;

            // 2. Normalize: strip diacritics + lowercase
            string norm = RemoveDiacriticsSimple(quan).ToLowerInvariant().Trim();
            // Bỏ prefix "quận ", "quan ", "huyện ", "tp ", "thành phố "
            norm = System.Text.RegularExpressions.Regex.Replace(
                norm, @"^(quan|huyen|tp|thanh pho)\s+", "");

            if (AppConstants.SHIPPING_FEES_BY_QUAN.TryGetValue(norm, out decimal fee2))
                return fee2;

            // 3. Thử chỉ lấy số (nếu quận số: "Quận 1" → "1")
            var numMatch = System.Text.RegularExpressions.Regex.Match(norm, @"\d+");
            if (numMatch.Success)
            {
                if (AppConstants.SHIPPING_FEES_BY_QUAN.TryGetValue(numMatch.Value, out decimal fee3))
                    return fee3;
            }

            return 0m;
        }

        /// <summary>
        /// Bỏ dấu tiếng Việt đơn giản (Bình Thạnh → Binh Thanh).
        /// Dùng nội bộ cho LookupShipFeeByQuan.
        /// </summary>
        private static string RemoveDiacriticsSimple(string s)
        {
            var norm = s.Normalize(System.Text.NormalizationForm.FormD);
            var sb = new System.Text.StringBuilder();
            foreach (char c in norm)
                if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c)
                    != System.Globalization.UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            return sb.ToString().Normalize(System.Text.NormalizationForm.FormC);
        }
    }
}

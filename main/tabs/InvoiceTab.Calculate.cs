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
                    colPhuong = -1,
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
                    if (header.Contains("phường") || header.Contains("phuong"))
                        colPhuong = col;
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
                decimal totalTienThuHangTon = 0,
                    totalTienShipHangTon = 0;
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

                    // Skip row "luu tra" / AT ngày cũ khỏi totalSoDon
                    // (giống logic bảng trái Excel: E cod = SUMIFS filter theo SHOP+NGÀY, loại "luu tra")
                    if (colNguoiDi >= 0 && colNguoiDi < row.Cells.Count)
                    {
                        string nguoiVal = (row.Cells[colNguoiDi].Value?.ToString() ?? "").Trim();
                        bool isLuuTra = AppConstants.NOT_SHIPPER_VALUES.Any(v =>
                            nguoiVal.Contains(v, StringComparison.OrdinalIgnoreCase)
                        );
                        bool isATOldDate =
                            nguoiVal.StartsWith(
                                AppConstants.NGUOI_DI_DEFAULT,
                                StringComparison.OrdinalIgnoreCase
                            )
                            && !nguoiVal.StartsWith(
                                AppConstants.NGUOI_DI_DEFAULT + DateTime.Now.ToString("dd-MM"),
                                StringComparison.OrdinalIgnoreCase
                            );
                        if (isLuuTra || isATOldDate)
                            continue;
                    }

                    // Detect hàng tồn (carry-over ngày trước)
                    bool isHangTon = false;
                    if (colHangTon >= 0 && colHangTon < row.Cells.Count)
                    {
                        string tonVal = (row.Cells[colHangTon].Value?.ToString() ?? "")
                            .Trim()
                            .ToLower();
                        isHangTon = tonVal == "x";
                    }

                    // Detect đơn trả cho LEFT summary: FAIL=xx AND ỨNG TIỀN=x → sum TIỀN HÀNG
                    bool isTraLeft = false;
                    if (
                        colFail >= 0
                        && colFail < row.Cells.Count
                        && colUngTien >= 0
                        && colUngTien < row.Cells.Count
                    )
                    {
                        string failVal = (row.Cells[colFail].Value?.ToString() ?? "")
                            .Trim()
                            .ToLower();
                        string ungVal = (row.Cells[colUngTien].Value?.ToString() ?? "")
                            .Trim()
                            .ToLower();
                        isTraLeft = failVal.Contains("xx") && ungVal == "x";
                    }
                    if (isTraLeft && colTienHang >= 0 && colTienHang < row.Cells.Count)
                    {
                        if (
                            decimal.TryParse(
                                row.Cells[colTienHang].Value?.ToString() ?? "",
                                out decimal hangVal
                            )
                        )
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
                var detailByNguoiDi = new Dictionary<string, NguoiDiDetail>(
                    StringComparer.OrdinalIgnoreCase
                );
                int totalDonGop = 0,
                    totalDonTra = 0;
                // Matches Excel COUNTIFS("*gộp*") — number of rows containing 'gộp'
                int totalGopCellCount = 0;
                // Number of rows that are both gộp AND đơn trả (COUNTIFS with both criteria)
                int totalDonTraGop = 0;
                // Number of rows marked 'hàng tỉnh' in ghi chú
                int totalDonHangTinh = 0;

                // Struct tạm collect row data per người đi (dùng cho gộp detection + đơn trả details)
                var rowsPerNguoi = new Dictionary<
                    string,
                    List<(
                        string TenKH,
                        string DiaChi,
                        string Quan,
                        string Phuong,
                        string Ma,
                        decimal TienThu,
                        decimal ShipFee,
                        bool IsTra,
                        string GhiChu
                    )>
                >(StringComparer.OrdinalIgnoreCase);

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

                        string nguoiRow =
                            colNguoiDi >= 0 && colNguoiDi < row.Cells.Count
                                ? (row.Cells[colNguoiDi].Value?.ToString() ?? "").Trim()
                                : "";
                        if (string.IsNullOrEmpty(nguoiRow))
                            nguoiRow = "(không rõ)";

                        // Không phải shipper thật (vd: "luu tra") → skip khỏi per-shipper summary
                        if (
                            AppConstants.NOT_SHIPPER_VALUES.Any(v =>
                                nguoiRow.Contains(v, StringComparison.OrdinalIgnoreCase)
                            )
                        )
                            continue;

                        // AT ngày cũ (VD: "AT 30-03" khi hôm nay 08-04) + FAIL="xx"
                        // → đơn trả thuộc AT hôm nay, sửa NGƯỜI ĐI thành "luu tra" để skip
                        bool isAnTamRow = nguoiRow.StartsWith(
                            AppConstants.NGUOI_DI_DEFAULT,
                            StringComparison.OrdinalIgnoreCase
                        );
                        string atToday =
                            AppConstants.NGUOI_DI_DEFAULT + DateTime.Now.ToString("dd-MM");
                        bool isAnTamOldDate =
                            isAnTamRow
                            && !nguoiRow.StartsWith(atToday, StringComparison.OrdinalIgnoreCase);
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

                        string tenKH =
                            colTenKH >= 0 && colTenKH < row.Cells.Count
                                ? (row.Cells[colTenKH].Value?.ToString() ?? "").Trim()
                                : "";
                        string diaChi =
                            colDiaChi >= 0 && colDiaChi < row.Cells.Count
                                ? (row.Cells[colDiaChi].Value?.ToString() ?? "").Trim()
                                : "";
                        string quan =
                            colQuan >= 0 && colQuan < row.Cells.Count
                                ? (row.Cells[colQuan].Value?.ToString() ?? "").Trim()
                                : "";
                        string phuong =
                            colPhuong >= 0 && colPhuong < row.Cells.Count
                                ? (row.Cells[colPhuong].Value?.ToString() ?? "").Trim()
                                : "";
                        string ma =
                            colMa >= 0 && colMa < row.Cells.Count
                                ? (row.Cells[colMa].Value?.ToString() ?? "").Trim()
                                : "";

                        // Detect đơn trả: GHI CHÚ contains "đơn trả"
                        // (trước dùng FAIL="xx" nhưng "xx" dễ trùng với giá trị khác)
                        bool isTra = false;
                        string ghiChuVal = "";
                        if (colGhiChu >= 0 && colGhiChu < row.Cells.Count)
                        {
                            ghiChuVal = (row.Cells[colGhiChu].Value?.ToString() ?? "")
                                .Trim()
                                .ToLower();
                            isTra = ghiChuVal.Contains("đơn trả");
                        }

                        // AT ngày cũ → đơn trả: tính tiền trừ vào AT hôm nay, rồi sửa thành "luu tra"
                        if (isAnTamOldDate)
                        {
                            // Đảm bảo AT hôm nay tồn tại trong detailByNguoiDi
                            if (!detailByNguoiDi.ContainsKey(atToday))
                                detailByNguoiDi[atToday] = new NguoiDiDetail();
                            var dAT = detailByNguoiDi[atToday];
                            dAT.IsAnTam = true;

                            // Tính đơn trả: -(tiền thu) dùng AT_SHIPPING_FEES
                            decimal shipFeeLookup = LookupShipFeeByDict(
                                quan,
                                AppConstants.AT_SHIPPING_FEES
                            );
                            decimal deduction = -(tienThuRow);
                            dAT.TienDonTra += deduction;
                            dAT.SoDonTra++;
                            dAT.DonTraDetails.Add((ma, tienThuRow, shipFeeLookup, deduction));

                            // Sửa NGƯỜI ĐI thành "luu tra" trong DataGridView
                            if (colNguoiDi >= 0 && colNguoiDi < row.Cells.Count)
                                row.Cells[colNguoiDi].Value = "luu tra";

                            // KHÔNG add vào rowsPerNguoi — deduction đã tính ở trên
                            continue; // Skip tích lũy bình thường
                        }

                        // Tích lũy per-person (dùng nguoiRow nguyên bản)
                        if (!detailByNguoiDi.ContainsKey(nguoiRow))
                            detailByNguoiDi[nguoiRow] = new NguoiDiDetail();
                        var d = detailByNguoiDi[nguoiRow];
                        d.TienThu += tienThuRow;
                        d.TienShip += tienShipRow;
                        d.SoDon++;
                        d.IsAnTam = nguoiRow.StartsWith(
                            atToday,
                            StringComparison.OrdinalIgnoreCase
                        );
                        if (isTra)
                            d.SoDonTra++;

                        // Collect row data cho gộp detection + đơn trả details
                        if (!rowsPerNguoi.ContainsKey(nguoiRow))
                            rowsPerNguoi[nguoiRow] =
                                new List<(
                                    string,
                                    string,
                                    string,
                                    string,
                                    string,
                                    decimal,
                                    decimal,
                                    bool,
                                    string
                                )>();
                        rowsPerNguoi[nguoiRow]
                            .Add(
                                (
                                    tenKH,
                                    diaChi,
                                    quan,
                                    phuong,
                                    ma,
                                    tienThuRow,
                                    tienShipRow,
                                    isTra,
                                    ghiChuVal
                                )
                            );
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
                    var groups = rows.Where(r =>
                            !string.IsNullOrEmpty(r.TenKH) && !string.IsNullOrEmpty(r.DiaChi)
                        )
                        .GroupBy(r => (r.TenKH.ToLower(), r.DiaChi.ToLower()));
                    // per-person counters for aggregated global totals
                    int gopCellsForPerson = 0;
                    int donTraGopForPerson = 0;
                    int hangTinhForPerson = 0;
                    foreach (var g in groups)
                    {
                        if (g.Count() > 1)
                        {
                            d.SoDonGop += g.Count() - 1;
                            gopCellsForPerson += g.Count();
                            // count how many rows in this group are marked 'đơn trả'
                            donTraGopForPerson += g.Count(r => r.IsTra);
                        }
                    }

                    // total đơn trả (all) kept for reporting, but money deduction uses only those that are both gộp & đơn trả
                    totalDonGop += d.SoDonGop;
                    totalDonTra += d.SoDonTra;

                    // count 'hàng tỉnh' flags in this person's rows
                    hangTinhForPerson = rows.Count(r =>
                        !string.IsNullOrEmpty(r.GhiChu) && r.GhiChu.Contains("hàng tỉnh")
                    );
                    totalGopCellCount += gopCellsForPerson;
                    totalDonTraGop += donTraGopForPerson;
                    totalDonHangTinh += hangTinhForPerson;

                    if (d.IsAnTam)
                    {
                        // AT: tính ship theo bảng phí AT zone (per ORDER, không trừ gộp)
                        var zoneBreakdown = new Dictionary<decimal, int>();
                        foreach (var r in rows)
                        {
                            decimal atFee = LookupShipFeeByDict(
                                r.Quan,
                                AppConstants.AT_SHIPPING_FEES
                            );
                            if (atFee == 0m)
                                continue; // quận không có trong bảng AT
                            if (!zoneBreakdown.ContainsKey(atFee))
                                zoneBreakdown[atFee] = 0;
                            zoneBreakdown[atFee]++;
                        }
                        d.AtZoneBreakdown = zoneBreakdown;
                        d.TienShipTru = 0;
                        foreach (var z in zoneBreakdown)
                            d.TienShipTru -= z.Key * z.Value;
                    }
                    else
                    {
                        // Shipper khác: công thức cũ -(TongShip - SoDonGiao × 5k)
                        d.TienShipTru = -(d.TienShip - d.SoDonGiao * AppConstants.PHI_SHIP_MOI_DON);
                    }

                    // Tiền lấy: không tính per-person, tính global cho NGUOI_LAY_DEFAULT
                    d.TienLay = 0;

                    // Đơn trả:
                    //   AT:     -(tienThu)  → trừ đúng tiền thu (hàng + ship AT), KHÔNG có phí công 5k
                    //   Khác:   -(tienThu - shipFee + 5k) → trừ tiền hàng + phí công 5k
                    // AT dùng AT_SHIPPING_FEES (zone riêng), shipper khác dùng SHIPPING_FEES_BY_QUAN
                    foreach (var r in rows.Where(r => r.IsTra))
                    {
                        decimal shipFeeLookup = d.IsAnTam
                            ? LookupShipFeeByDict(r.Quan, AppConstants.AT_SHIPPING_FEES)
                            : LookupShipFee(r.Phuong, r.Quan);
                        decimal deduction = d.IsAnTam
                            ? -(r.TienThu) // AT: trả lại toàn bộ tiền thu (hàng + ship AT)
                            : -(r.TienThu - shipFeeLookup + AppConstants.PHI_CONG_DON_TRA);
                        d.TienDonTra += deduction;
                        d.DonTraDetails.Add((r.Ma, r.TienThu, shipFeeLookup, deduction));
                    }
                }

                // Tính tiền lấy global cho NGUOI_LAY_DEFAULT (c.cuong)
                // Use same rules as Excel formula (dynamic):
                // donLay = totalSoDon - (COUNTIFS("*gộp*")/2) - (COUNTIFS(gộp & đơn trả)/2) - COUNTIFS("*hàng tỉnh*")
                decimal donLayGlobal =
                    totalSoDon
                    - ((decimal)totalGopCellCount / 2m)
                    - ((decimal)totalDonTraGop / 2m)
                    - totalDonHangTinh;
                if (donLayGlobal < 0)
                    donLayGlobal = 0;
                decimal tienLayTong = -(donLayGlobal * AppConstants.PHI_LAY_HANG_MOI_DON);

                // Gán tiền lấy vào đúng người lấy (c.cuong) nếu có trong detailByNguoiDi
                foreach (var kvp in detailByNguoiDi)
                {
                    if (
                        kvp.Key.Equals(
                            AppConstants.NGUOI_LAY_DEFAULT,
                            StringComparison.OrdinalIgnoreCase
                        )
                    )
                    {
                        kvp.Value.TienLay = tienLayTong;
                        break;
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
                    TienLayTong = tienLayTong,
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
        /// Tra cứu phí ship theo PHƯỜNG + QUẬN.
        /// Thứ tự ưu tiên:
        ///   1. Composite key "quận:phường" trong SHIPPING_FEES_BY_WARD  (VD: "8:5" → 30k)
        ///   2. Plain ward name trong SHIPPING_FEES_BY_WARD               (VD: "rach ong" → 25k)
        ///   3. Fallback SHIPPING_FEES_BY_QUAN                           (VD: "8" → 30k)
        /// </summary>
        private static decimal LookupShipFee(string phuong, string quan)
        {
            if (!string.IsNullOrWhiteSpace(phuong))
            {
                // Normalize phường: bỏ dấu + lowercase + strip prefix "phường/phuong/p."
                string normP = RemoveDiacriticsSimple(phuong).ToLowerInvariant().Trim();
                normP = System
                    .Text.RegularExpressions.Regex.Replace(normP, @"^(phuong|p\.?)\s*", "")
                    .Trim();

                // Normalize quận → số hoặc tên không dấu (giống LookupShipFeeByDict)
                string normQ = RemoveDiacriticsSimple(quan ?? "").ToLowerInvariant().Trim();
                normQ = System.Text.RegularExpressions.Regex.Replace(
                    normQ,
                    @"^(quan|huyen|tp|thanh pho)\s+",
                    ""
                );
                var qNum = System.Text.RegularExpressions.Regex.Match(normQ, @"\d+");
                if (qNum.Success)
                    normQ = qNum.Value;

                // 1. Composite "quan:phuong" — handles Q8 old numbers & named wards
                if (!string.IsNullOrEmpty(normQ) && !string.IsNullOrEmpty(normP))
                    if (
                        AppConstants.SHIPPING_FEES_BY_WARD.TryGetValue(
                            $"{normQ}:{normP}",
                            out decimal feeC
                        )
                    )
                        return feeC;

                // 2. Plain ward name — handles new named wards (rach ong, hung phu, …)
                if (!string.IsNullOrEmpty(normP))
                    if (AppConstants.SHIPPING_FEES_BY_WARD.TryGetValue(normP, out decimal feeW))
                        return feeW;
            }

            // 3. Fallback to district fee
            return LookupShipFeeByQuan(quan);
        }

        /// <summary>
        /// Tra cứu phí ship theo QUẬN từ AppConstants.SHIPPING_FEES_BY_QUAN.
        /// Tự bỏ dấu tiếng Việt + normalize trước khi tra.
        /// Trả về 0 nếu không tìm thấy.
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

            // 1. Thử exact match (case-insensitive đã có trong dictionary)
            if (feeDict.TryGetValue(quan.Trim(), out decimal fee1))
                return fee1;

            // 2. Normalize: strip diacritics + lowercase
            string norm = RemoveDiacriticsSimple(quan).ToLowerInvariant().Trim();
            // Bỏ prefix "quận ", "quan ", "huyện ", "tp ", "thành phố "
            norm = System.Text.RegularExpressions.Regex.Replace(
                norm,
                @"^(quan|huyen|tp|thanh pho)\s+",
                ""
            );

            if (feeDict.TryGetValue(norm, out decimal fee2))
                return fee2;

            // 3. Thử chỉ lấy số (nếu quận số: "Quận 1" → "1")
            var numMatch = System.Text.RegularExpressions.Regex.Match(norm, @"\d+");
            if (numMatch.Success)
            {
                if (feeDict.TryGetValue(numMatch.Value, out decimal fee3))
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
            s = s.Replace('đ', 'd').Replace('Đ', 'D');
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
    }
}

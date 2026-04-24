using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace TextInputter
{
    // ─── Daily Report Display + Save ────────────────────────────────────────────
    public partial class MainForm
    {
        private void DisplayDailyReport()
        {
            if (currentDailyReport == null)
                return;

            Panel pnlTop = tabInvoice.Controls["pnlInvoiceTop"] as Panel;
            Panel pnlBottom = tabInvoice.Controls["pnlDailyReportBottom"] as Panel;

            if (pnlTop == null)
            {
                // First time: tạo layout từ đầu

                // Bottom panel: report panels (AutoScroll để kéo ngang)
                pnlBottom = new Panel
                {
                    Name = "pnlDailyReportBottom",
                    Dock = DockStyle.Bottom,
                    Height = AppConstants.DAILY_REPORT_PANEL_HEIGHT,
                    BackColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    AutoScroll = true,
                };

                // Rebuild layout: cùng pattern với InitializeInvoiceTabUI —
                // Fill add TRƯỚC, Top/Bottom add SAU.
                // Đảm bảo dgvInvoice lấy phần còn lại sau khi Bottom/Top đã dock.
                tabInvoice.Controls.Remove(dgvInvoice);
                tabInvoice.Controls.Remove(lblInvoiceTotal);
                // Add lại theo thứ tự chuẩn
                tabInvoice.Controls.Add(dgvInvoice); // Fill — add trước
                tabInvoice.Controls.Add(lblInvoiceTotal); // Top  — add sau
                tabInvoice.Controls.Add(pnlBottom); // Bottom — add sau

                // pnlTop dùng để check "đã khởi tạo" — tạo dummy để các lần sau skip
                pnlTop = new Panel
                {
                    Name = "pnlInvoiceTop",
                    Visible = false,
                    Width = 0,
                    Height = 0,
                };
                tabInvoice.Controls.Add(pnlTop);
            }

            pnlBottom.Controls.Clear();

            var r = currentDailyReport;
            // LEFT summary: loại hàng tồn (carry-over ngày trước, HÀNG TỒN=x)
            decimal thuLeft = r.TongTienThu - r.TongTienThuHangTon;
            decimal shipLeft = r.TongTienShip - r.TongTienShipHangTon;
            decimal hangLeft = thuLeft - shipLeft; // TIỀN HÀNG = TIỀN THU - TIỀN SHIP (loại hàng tồn)
            decimal soDonLeft = r.SoDon - r.SoDonHangTon;
            string soDonStr = soDonLeft.ToString("N0");
            string hangStr = hangLeft.ToString("N0");
            decimal tongShipRaw = 0m; // Ship đã trừ sẵn trong TIỀN HÀNG
            // KẾT = dùng TongKetCuoi đã tính ở BtnCalculateExcelData_Click (bao gồm cả đơn âm)
            decimal ketTong = r.TongKetCuoi;
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
                g.Columns[1].Width = 120;
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
            int nguoiPanelWidth = 360;
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
                // LEFT "Đơn trả": -SUMIFS(TIỀN HÀNG, FAIL="xx", ỨNG TIỀN="x")
                // matching Excel formula (not per-person deduction)
                decimal tienHangDonTra = -r.TongTienHangDonTra; // negative sum of TIỀN HÀNG

                ri = dgvTong.Rows.Add("", "Tiền", "Số đơn");
                dgvTong.Rows[ri].DefaultCellStyle.BackColor = Color.LightSteelBlue;
                dgvTong.Rows[ri].DefaultCellStyle.Font = new Font("Arial", 10, FontStyle.Bold);

                // Tiền hàng (= tổng TIỀN HÀNG từ khách — matching template "cod")
                ri = dgvTong.Rows.Add("Tiền hàng", hangStr, soDonStr);
                dgvTong.Rows[ri].DefaultCellStyle.BackColor = Color.White;

                // Trừ Ship (placeholder = 0, ship đã trừ sẵn trong Tiền hàng)
                ri = dgvTong.Rows.Add("Trừ Ship", tongShipRaw.ToString("N0"), "");
                dgvTong.Rows[ri].DefaultCellStyle.ForeColor = Color.Red;

                // Đơn trả & c.khoản — matching Excel: -SUMIFS(TIỀN HÀNG, FAIL="xx", ỨNG TIỀN="x")
                if (r.SoDonTraLeft > 0)
                {
                    ri = dgvTong.Rows.Add(
                        "Đơn trả & c.khoản",
                        tienHangDonTra.ToString("N0"),
                        r.SoDonTraLeft.ToString()
                    );
                    dgvTong.Rows[ri].DefaultCellStyle.ForeColor = Color.Red;
                }
                else
                {
                    ri = dgvTong.Rows.Add("Đơn trả & c.khoản", "", "");
                    dgvTong.Rows[ri].DefaultCellStyle.ForeColor = Color.Gray;
                }

                // Tiền Hàng = Tiền thu + Trừ Ship + Đơn trả (matching manual "Tiền Hàng Hcm")
                decimal tienHangFinal = thuLeft + tongShipRaw + tienHangDonTra;
                ri = dgvTong.Rows.Add("Tiền Hàng", tienHangFinal.ToString("N0"), soDonStr);
                dgvTong.Rows[ri].DefaultCellStyle.BackColor = Color.FromArgb(230, 245, 255);
                dgvTong.Rows[ri].DefaultCellStyle.Font = new Font("Arial", 10, FontStyle.Bold);

                // đền đơn (placeholder — user tự điền)
                ri = dgvTong.Rows.Add("đền đơn", "", "");
                dgvTong.Rows[ri].DefaultCellStyle.ForeColor = Color.Red;

                // nợ cũ — auto-filled từ NegativeRows nếu có
                decimal totalNegative = r.NegativeRows?.Sum(nr => nr.Amount) ?? 0;
                if (r.NegativeRows != null && r.NegativeRows.Count > 0)
                {
                    string negLabel = r.NegativeRows.Count == 1 ? r.NegativeRows[0].Label : "nợ cũ";
                    ri = dgvTong.Rows.Add(
                        negLabel,
                        totalNegative.ToString("N0"),
                        $"{r.NegativeRows.Count} dòng"
                    );
                    dgvTong.Rows[ri].DefaultCellStyle.ForeColor = Color.FromArgb(200, 100, 0); // cam
                }
                else
                {
                    ri = dgvTong.Rows.Add("nợ cũ", "", "");
                    dgvTong.Rows[ri].DefaultCellStyle.ForeColor = Color.FromArgb(200, 100, 0);
                }

                // THANH TOÁN = Tiền Hàng + đền đơn + nợ cũ
                decimal tongFinal = tienHangFinal + totalNegative;
                ri = dgvTong.Rows.Add("THANH TOÁN", tongFinal.ToString("N0"), soDonStr);
                dgvTong.Rows[ri].DefaultCellStyle.BackColor = Color.LightGreen;
                dgvTong.Rows[ri].DefaultCellStyle.Font = new Font("Arial", 11, FontStyle.Bold);
                dgvTong.Rows[ri].Height = AppConstants.ROW_HEIGHT_REPORT_KET;

                pnlTong.Controls.Add(dgvTong);
                pnlReports.Controls.Add(pnlTong);
            }

            // ── Report nhỏ theo từng NGƯỜI ĐI ────────────────────────────────
            if (r.DetailByNguoiDi != null && r.DetailByNguoiDi.Count > 0)
            {
                foreach (var kvp in r.DetailByNguoiDi.OrderBy(k => k.Key))
                {
                    string tenNguoi = kvp.Key;
                    var nd = kvp.Value;
                    decimal tienThuNguoi = nd.TienThu;
                    decimal tienShipNguoi = nd.TienShip;
                    // Show delivered orders = total receipts - grouped duplicates
                    decimal soDonNguoi = nd.SoDon - nd.SoDonGop;

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
                        "TỔNG ĐƠN NHẬN",
                        tienThuNguoi.ToString("N0"),
                        soDonNguoi.ToString("N0")
                    );
                    dgvNguoi.Rows[ri].DefaultCellStyle.BackColor = Color.White;

                    if (nd.IsAnTam && nd.AtZoneBreakdown.Count > 0)
                    {
                        // AT: hiển thị zone breakdown trước tiền ship tổng
                        foreach (var zone in nd.AtZoneBreakdown.OrderBy(z => z.Key))
                        {
                            decimal zoneTotal = -(zone.Key * zone.Value);
                            ri = dgvNguoi.Rows.Add(
                                $"  {zone.Value} × {zone.Key:N0}k",
                                zoneTotal.ToString("N0"),
                                ""
                            );
                            dgvNguoi.Rows[ri].DefaultCellStyle.ForeColor = Color.Gray;
                            dgvNguoi.Rows[ri].DefaultCellStyle.BackColor = Color.FromArgb(
                                245,
                                245,
                                245
                            );
                        }
                        // tiền ship tổng cho AT
                        ri = dgvNguoi.Rows.Add(
                            "tiền ship",
                            nd.TienShipTru.ToString("N0"),
                            $"{soDonNguoi:N0} đơn"
                        );
                        dgvNguoi.Rows[ri].DefaultCellStyle.BackColor = Color.White;
                        if (nd.TienShipTru < 0)
                            dgvNguoi.Rows[ri].Cells[1].Style.ForeColor = Color.Red;
                    }
                    else
                    {
                        // Shipper khác: công thức cũ -(TongShip - SoDonGiao × 5k)
                        string shipInfo =
                            nd.SoDonGop > 0
                                ? $"giao {nd.SoDonGiao:N0} ({nd.SoDonGop} gộp)"
                                : $"giao {nd.SoDonGiao:N0}";
                        ri = dgvNguoi.Rows.Add(
                            "tiền ship",
                            nd.TienShipTru.ToString("N0"),
                            shipInfo
                        );
                        dgvNguoi.Rows[ri].DefaultCellStyle.BackColor = Color.White;
                        if (nd.TienShipTru < 0)
                            dgvNguoi.Rows[ri].Cells[1].Style.ForeColor = Color.Red;
                    }

                    // tiền lấy — chỉ hiện giá trị cho NGUOI_LAY_DEFAULT (c.cuong)
                    bool isNguoiLay = tenNguoi.Equals(
                        AppConstants.NGUOI_LAY_DEFAULT,
                        StringComparison.OrdinalIgnoreCase
                    );
                    if (isNguoiLay && nd.TienLay != 0)
                    {
                        decimal donLayCount = r.SoDon - r.TotalDonGop - r.TotalDonTra;
                        if (donLayCount < 0)
                            donLayCount = 0;
                        ri = dgvNguoi.Rows.Add(
                            "tiền lấy",
                            nd.TienLay.ToString("N0"),
                            $"{donLayCount:N0}"
                        );
                        dgvNguoi.Rows[ri].DefaultCellStyle.BackColor = Color.White;
                        if (nd.TienLay < 0)
                            dgvNguoi.Rows[ri].Cells[1].Style.ForeColor = Color.Red;
                    }
                    else
                    {
                        ri = dgvNguoi.Rows.Add("tiền lấy", "", "");
                        dgvNguoi.Rows[ri].DefaultCellStyle.BackColor = Color.White;
                    }

                    // đơn trả — auto-filled nếu có, đỏ nếu chưa có
                    if (nd.SoDonTra > 0)
                    {
                        ri = dgvNguoi.Rows.Add(
                            "đơn trả",
                            nd.TienDonTra.ToString("N0"),
                            $"{nd.SoDonTra} đơn"
                        );
                        dgvNguoi.Rows[ri].DefaultCellStyle.ForeColor = Color.Red;
                    }
                    else
                    {
                        ri = dgvNguoi.Rows.Add("đơn trả", "0", "0");
                        dgvNguoi.Rows[ri].DefaultCellStyle.ForeColor = Color.Gray;
                    }

                    // đơn cũ ck (placeholder đỏ, vẫn cần user tự điền)
                    ri = dgvNguoi.Rows.Add("đơn cũ ck", "", "");
                    dgvNguoi.Rows[ri].DefaultCellStyle.ForeColor = Color.Red;

                    // Dòng KẾT — tính tự động: thu + ship + lấy + trả
                    decimal ketNguoi = tienThuNguoi + nd.TienShipTru + nd.TienLay + nd.TienDonTra;
                    ri = dgvNguoi.Rows.Add(
                        "KẾT",
                        ketNguoi.ToString("N0"),
                        soDonNguoi.ToString("N0")
                    );
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

            // WinForms Dock z-order rule:
            //   Docking xử lý từ control add SAU (z-front) đến control add TRƯỚC (z-back).
            //   → Fill phải add TRƯỚC (z-back), Top/Bottom add SAU (z-front).
            // Thứ tự visual mong muốn: pnlButtons (Top) → lblInvoiceTotal (Top) → dgvInvoice (Fill) → pnlBottom (Bottom)
            tabInvoice.Controls.Remove(dgvInvoice);
            tabInvoice.Controls.Remove(lblInvoiceTotal);
            Panel existingBottom = tabInvoice.Controls["pnlDailyReportBottom"] as Panel;
            if (existingBottom != null)
                tabInvoice.Controls.Remove(existingBottom);

            tabInvoice.Controls.Add(dgvInvoice); // Fill — add trước (z-back)
            if (existingBottom != null)
                tabInvoice.Controls.Add(existingBottom); // Bottom — add sau
            tabInvoice.Controls.Add(lblInvoiceTotal); // Top  — add sau
            tabInvoice.Controls.Add(pnlButtons); // Top  — add cuối (z-front → dock trên cùng)

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
            btnSave.Click += async (s, e) => await SaveDailyReportToExcelAsync(btnSave);
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

        private async Task SaveDailyReportToExcelAsync(Button callerBtn = null)
        {
            if (callerBtn != null)
                callerBtn.Enabled = false;
            Panel overlay = null;
            try
            {
                if (dgvInvoice.Rows.Count == 0)
                {
                    MessageBox.Show(
                        "Không có dữ liệu để lưu!",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }

                // ── Chọn file lưu (thường là file tháng hiện tại) ────────────
                string defaultName =
                    (currentDailyReport?.Date is string d2 && !string.IsNullOrEmpty(d2))
                        ? $"DailyReport_{d2.Replace("/", "-").Replace(".", "-")}.xlsx"
                        : "DailyReport.xlsx";

                using var sfd = new SaveFileDialog
                {
                    Title = "Lưu báo cáo hàng ngày",
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    DefaultExt = "xlsx",
                    FileName = defaultName,
                    InitialDirectory = System.IO.Directory.Exists(
                        System.IO.Path.GetDirectoryName(currentExcelFilePath ?? "")
                    )
                        ? System.IO.Path.GetDirectoryName(currentExcelFilePath)
                        : AppDomain.CurrentDomain.BaseDirectory,
                };
                if (sfd.ShowDialog() != DialogResult.OK)
                    return;
                string excelPath = sfd.FileName;

                // ── Nếu file có sẵn → hỏi sheet name ─────────────────────────
                string sheetName;
                if (System.IO.File.Exists(excelPath))
                {
                    List<string> existingSheets;
                    try
                    {
                        using var wbRead = new XLWorkbook(excelPath);
                        existingSheets = wbRead.Worksheets.Select(ws => ws.Name).ToList();
                    }
                    catch
                    {
                        existingSheets = new List<string>();
                    }

                    string suggested =
                        tabExcelSheets.SelectedTab?.Text ?? DateTime.Now.ToString("dd-MM");
                    sheetName = PickOrCreateSheetName(existingSheets, suggested);
                    if (sheetName == null)
                        return; // user hủy
                }
                else
                {
                    sheetName = tabExcelSheets.SelectedTab?.Text ?? DateTime.Now.ToString("dd-MM");
                }

                // ── Snapshot data từ UI thread ────────────────────────────────
                // Chỉ lấy các data row thực sự — bỏ qua row ▶ TỔNG, ▶ KẾT (row tổng kết do UI tạo ra)
                var dgvRows = dgvInvoice
                    .Rows.Cast<DataGridViewRow>()
                    .Where(r =>
                    {
                        if (r.IsNewRow)
                            return false;
                        // Bỏ row tổng kết (cell 0 bắt đầu bằng "▶")
                        string firstCell =
                            r.Cells.Count > 0 ? r.Cells[0].Value?.ToString() ?? "" : "";
                        if (firstCell.StartsWith("▶"))
                            return false;
                        return true;
                    })
                    .Select(r =>
                        r.Cells.Cast<DataGridViewCell>()
                            .Select(c => c.Value?.ToString() ?? "")
                            .ToList()
                    )
                    .ToList();
                var colHeaders = dgvInvoice
                    .Columns.Cast<DataGridViewColumn>()
                    .Select(c => c.HeaderText)
                    .ToList();

                // ── Ghi workbook (background thread) ──────────────────────────
                overlay = ShowLoadingOverlay("⏳ Đang ghi dữ liệu vào Excel...");
                await Task.Run(() =>
                    WriteSheetToWorkbook(excelPath, sheetName, colHeaders, dgvRows)
                );

                // ── Ghi formula + bảng tổng kết ───────────────────────────────
                HideLoadingOverlay(overlay);
                overlay = ShowLoadingOverlay("⏳ Đang ghi công thức + bảng tổng kết...");

                DateTime sheetDate = DateTime.Now;
                if (
                    !DateTime.TryParseExact(
                        sheetName,
                        "dd-MM",
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None,
                        out sheetDate
                    )
                )
                    sheetDate = DateTime.Now;

                var svc = new TextInputter.Services.ExcelInvoiceService(excelPath);
                await Task.Run(() => svc.ApplyFormulasAndSummary(sheetName, sheetDate));

                MessageBox.Show(
                    $"✅ Đã lưu thành công!\nFile: {System.IO.Path.GetFileName(excelPath)}\nSheet: {sheetName}",
                    "✅ Lưu thành công",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                lblStatus.Text = $"✅ {System.IO.Path.GetFileName(excelPath)} [{sheetName}]";
                lblStatus.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌ Lỗi khi lưu: {ex.Message}",
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                Debug.WriteLine($"SaveDailyReport error: {ex}");
            }
            finally
            {
                if (overlay != null)
                    HideLoadingOverlay(overlay);
                if (callerBtn != null)
                    callerBtn.Enabled = true;
            }
        }

        /// <summary>
        /// Dialog cho user chọn sheet có sẵn hoặc nhập tên sheet mới.
        /// Trả về null nếu user nhấn Hủy.
        /// </summary>
        private string PickOrCreateSheetName(List<string> existingSheets, string suggested)
        {
            using var dlg = new Form
            {
                Text = "Chọn hoặc tạo sheet",
                Size = new System.Drawing.Size(390, 250),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
            };

            // Radio: chọn sheet có sẵn
            var rbExisting = new RadioButton
            {
                Text = "Ghi đè sheet có sẵn:",
                Left = 12,
                Top = 12,
                Width = 340,
                Checked = existingSheets.Count > 0,
            };
            var cmb = new ComboBox
            {
                Left = 12,
                Top = 34,
                Width = 354,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Enabled = existingSheets.Count > 0,
            };
            if (existingSheets.Count > 0)
            {
                cmb.Items.AddRange(existingSheets.Cast<object>().ToArray());
                int idx = existingSheets.IndexOf(suggested);
                cmb.SelectedIndex = idx >= 0 ? idx : 0;
            }

            // Radio: tạo sheet mới
            var rbNew = new RadioButton
            {
                Text = "Tạo sheet mới:",
                Left = 12,
                Top = 72,
                Width = 340,
                Checked = existingSheets.Count == 0,
            };
            var txt = new TextBox
            {
                Left = 12,
                Top = 94,
                Width = 354,
                Text = suggested,
                Enabled = existingSheets.Count == 0,
            };

            rbExisting.CheckedChanged += (_, __) =>
            {
                cmb.Enabled = rbExisting.Checked;
                txt.Enabled = !rbExisting.Checked;
            };
            rbNew.CheckedChanged += (_, __) =>
            {
                cmb.Enabled = !rbNew.Checked;
                txt.Enabled = rbNew.Checked;
            };

            var btnOk = new Button
            {
                Text = "OK",
                Left = 214,
                Top = 170,
                Width = 75,
                DialogResult = DialogResult.OK,
            };
            var btnCancel = new Button
            {
                Text = "Hủy",
                Left = 298,
                Top = 170,
                Width = 75,
                DialogResult = DialogResult.Cancel,
            };

            dlg.Controls.AddRange(new Control[] { rbExisting, cmb, rbNew, txt, btnOk, btnCancel });
            dlg.AcceptButton = btnOk;
            dlg.CancelButton = btnCancel;

            if (dlg.ShowDialog() != DialogResult.OK)
                return null;

            return rbNew.Checked
                ? (txt.Text.Trim().Length > 0 ? txt.Text.Trim() : suggested)
                : (cmb.SelectedItem?.ToString() ?? suggested);
        }

        /// <summary>
        /// Mở workbook (hoặc tạo mới), xóa sheet cũ nếu trùng tên,
        /// ghi header + data, rồi SaveAs — KHÔNG đụng đến các sheet khác.
        /// </summary>
        private static void WriteSheetToWorkbook(
            string excelPath,
            string sheetName,
            List<string> colHeaders,
            List<List<string>> dgvRows
        )
        {
            // Mở file có sẵn → giữ nguyên tất cả sheet khác
            // Tạo file mới nếu chưa tồn tại
            var workbook = System.IO.File.Exists(excelPath)
                ? new XLWorkbook(excelPath)
                : new XLWorkbook();

            using (workbook)
            {
                // Xóa sheet cũ cùng tên (nếu có) trước khi tạo lại
                if (workbook.TryGetWorksheet(sheetName, out _))
                    workbook.Worksheets.Delete(sheetName);

                var ws = workbook.Worksheets.Add(sheetName);

                // Row 2: tiêu đề cột (matching DATA_START_ROW = 3, header at row 2)
                for (int c = 0; c < colHeaders.Count; c++)
                {
                    var cell = ws.Cell(2, c + 1);
                    cell.Value = colHeaders[c];
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                }

                // Row 3 trở đi: data (row 2 để trống cho cấu trúc DATA_START_ROW=3)
                for (int r = 0; r < dgvRows.Count; r++)
                {
                    for (int c = 0; c < dgvRows[r].Count; c++)
                    {
                        var cell = ws.Cell(r + 3, c + 1);
                        string val = dgvRows[r][c];
                        if (
                            double.TryParse(
                                val,
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture,
                                out double num
                            )
                        )
                        {
                            cell.Value = num;
                            cell.Style.NumberFormat.Format = "#,##0";
                        }
                        else
                            cell.Value = val;
                    }
                }

                ws.Columns().AdjustToContents();
                workbook.SaveAs(excelPath);
            }
        }
    }
}

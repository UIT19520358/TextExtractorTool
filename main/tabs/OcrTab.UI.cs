using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace TextInputter
{
    /// <summary>
    /// OcrTab UI — control field declarations + InitializeOCRTab().
    /// Logic (ProcessImages, SelectOCRFolder, ExportMappedDataToExcel...) ở OcrTab.cs.
    /// </summary>
    public partial class MainForm
    {
        // ─── Controls thuộc OCR Tab ────────────────────────────────────────────
        private ComboBox txtNguoiDiOCR;
        private ComboBox txtNguoiLayOCR;
        private RichTextBox txtRawOCRLog;
        private RichTextBox txtProcessLog;

        /// <summary>
        /// true  = người dùng tự nhập Người Đi (bỏ qua auto-map theo quận).
        /// false = auto-map theo AREA_TO_NGUOI_DI (logic cũ).
        /// </summary>
        private bool _manualNguoiDi = false;

        /// <summary>
        /// true  = người dùng tự nhập Người Lấy.
        /// false = dùng NGUOI_LAY_DEFAULT.
        /// </summary>
        private bool _manualNguoiLay = false;

        /// <summary>
        /// true  = dùng ngày hôm nay làm sheet name (tất cả vào 1 sheet).
        /// false = group theo NGÀY LẤY trong từng hóa đơn (mỗi ngày 1 sheet).
        /// null  = dùng ngày tự nhập (_exportCustomDate).
        /// </summary>
        private bool? _exportUseToday = false;

        /// <summary>
        /// Ngày tự nhập khi chọn "Ngày khác", format "dd-MM" (VD: "25-02").
        /// Chỉ dùng khi _exportUseToday == null.
        /// </summary>
        private string _exportCustomDate = "";

        // ─── Init ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Khởi tạo tab OCR: folder selection, người đi/lấy, raw log, mapping log, export button.
        /// Gọi từ MainForm constructor sau InitializeComponent().
        /// </summary>
        private void InitializeOCRTab()
        {
            try
            {
                Panel pnlOCR = new Panel
                {
                    Dock = DockStyle.Fill,
                    AutoScroll = true,
                    BackColor = SystemColors.Control,
                    Padding = new Padding(10),
                };

                int y = 10;

                // ── Title ──────────────────────────────────────────────────────
                UIHelper.CreateSectionLabel(pnlOCR, "🔍 OCR Processing", ref y);
                y -= 15;

                // ── Folder selection ───────────────────────────────────────────
                Label lblFolderInfo = new Label
                {
                    Text = "Chọn folder ảnh để quét OCR tự động",
                    AutoSize = true,
                    Location = new Point(10, y),
                    Font = new Font("Arial", 10, FontStyle.Bold),
                };
                pnlOCR.Controls.Add(lblFolderInfo);
                y += 25;

                var btnSelectFolder = UIHelper.CreateButton(
                    "📂 Chọn Folder",
                    Color.LightBlue,
                    10,
                    y,
                    130,
                    35
                );
                btnSelectFolder.Click += (s, e) => SelectOCRFolder();
                pnlOCR.Controls.Add(btnSelectFolder);

                var btnStartScan = UIHelper.CreateButton(
                    "▶ Bắt Đầu Quét",
                    Color.LightGreen,
                    150,
                    y,
                    130,
                    35
                );
                btnStartScan.Click += (s, e) => btnStart_Click(null, EventArgs.Empty);
                pnlOCR.Controls.Add(btnStartScan);

                var btnLoadFromLog = UIHelper.CreateButton(
                    "📋 Tải từ Log",
                    Color.LightYellow,
                    290,
                    y,
                    130,
                    35
                );
                btnLoadFromLog.Click += (s, e) => LoadFromLog();
                new ToolTip().SetToolTip(
                    btnLoadFromLog,
                    "Đọc ocr_log.txt từ lần quét trước → khôi phục dữ liệu, không tốn quota API"
                );
                pnlOCR.Controls.Add(btnLoadFromLog);
                y += 45;

                // ── Người Đi ──────────────────────────────────────────────────
                UIHelper.CreateSectionLabel(pnlOCR, "Thông tin NGƯỜI ĐI & NGƯỜI LẤY:", ref y);
                y -= 15;

                pnlOCR.Controls.Add(
                    new Label
                    {
                        Text = "Người Đi:",
                        AutoSize = true,
                        Location = new Point(10, y),
                        Font = new Font("Arial", 9, FontStyle.Bold),
                    }
                );
                y += 20;

                var rdoNguoiDiAuto = new RadioButton
                {
                    Text = "⚙ Tự động theo quận",
                    Location = new Point(10, y),
                    AutoSize = true,
                    Font = new Font("Arial", 9),
                    Checked = true,
                };
                var rdoNguoiDiManual = new RadioButton
                {
                    Text = "✏ Tự nhập:",
                    Location = new Point(175, y),
                    AutoSize = true,
                    Font = new Font("Arial", 9),
                };
                txtNguoiDiOCR = new ComboBox
                {
                    Location = new Point(265, y - 1),
                    Width = 140,
                    Font = new Font("Arial", 9),
                    DropDownStyle = ComboBoxStyle.DropDown,
                    Text = AppConstants.NGUOI_DI_DEFAULT  + DateTime.Now.ToString("dd-MM"),
                    Enabled = false,
                };
                txtNguoiDiOCR.Items.AddRange(AppConstants.NGUOI_LIST);

                rdoNguoiDiAuto.CheckedChanged += (s, e) =>
                {
                    if (rdoNguoiDiAuto.Checked)
                    {
                        _manualNguoiDi = false;
                        txtNguoiDiOCR.Enabled = false;
                    }
                };
                rdoNguoiDiManual.CheckedChanged += (s, e) =>
                {
                    if (rdoNguoiDiManual.Checked)
                    {
                        _manualNguoiDi = true;
                        txtNguoiDiOCR.Enabled = true;
                        txtNguoiDiOCR.Focus();
                    }
                };

                pnlOCR.Controls.Add(rdoNguoiDiAuto);
                pnlOCR.Controls.Add(rdoNguoiDiManual);
                pnlOCR.Controls.Add(txtNguoiDiOCR);
                y += 32;

                // ── Người Lấy ─────────────────────────────────────────────────
                pnlOCR.Controls.Add(
                    new Label
                    {
                        Text = "Người Lấy:",
                        AutoSize = true,
                        Location = new Point(10, y),
                        Font = new Font("Arial", 9, FontStyle.Bold),
                    }
                );
                y += 20;

                var rdoNguoiLayAuto = new RadioButton
                {
                    Text = "⚙ Tự động (" + AppConstants.NGUOI_LAY_DEFAULT + ")",
                    Location = new Point(10, y),
                    AutoSize = true,
                    Font = new Font("Arial", 9),
                    Checked = true,
                };
                var rdoNguoiLayManual = new RadioButton
                {
                    Text = "✏ Tự nhập:",
                    Location = new Point(175, y),
                    AutoSize = true,
                    Font = new Font("Arial", 9),
                };
                txtNguoiLayOCR = new ComboBox
                {
                    Location = new Point(265, y - 1),
                    Width = 140,
                    Font = new Font("Arial", 9),
                    DropDownStyle = ComboBoxStyle.DropDown,
                    Text = AppConstants.NGUOI_LAY_DEFAULT,
                    Enabled = false,
                };
                txtNguoiLayOCR.Items.AddRange(AppConstants.NGUOI_LIST);

                rdoNguoiLayAuto.CheckedChanged += (s, e) =>
                {
                    if (rdoNguoiLayAuto.Checked)
                    {
                        _manualNguoiLay = false;
                        txtNguoiLayOCR.Enabled = false;
                    }
                };
                rdoNguoiLayManual.CheckedChanged += (s, e) =>
                {
                    if (rdoNguoiLayManual.Checked)
                    {
                        _manualNguoiLay = true;
                        txtNguoiLayOCR.Enabled = true;
                        txtNguoiLayOCR.Focus();
                    }
                };

                pnlOCR.Controls.Add(rdoNguoiLayAuto);
                pnlOCR.Controls.Add(rdoNguoiLayManual);
                pnlOCR.Controls.Add(txtNguoiLayOCR);
                y += 36;

                // ── Raw OCR Log ───────────────────────────────────────────────
                UIHelper.CreateSectionLabel(pnlOCR, "📋 Raw OCR Text (Kết quả OCR thô):", ref y);
                y -= 15;

                var rawSearchPanel = CreateSearchBarForRaw(pnlOCR, y);
                y += 32;

                txtRawOCRLog = new RichTextBox
                {
                    Location = new Point(10, y),
                    Width = this.ClientSize.Width - 50,
                    Height = 200,
                    ReadOnly = true,
                    BackColor = Color.White,
                    Font = new Font("Courier New", 8),
                    BorderStyle = BorderStyle.FixedSingle,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                };
                pnlOCR.Controls.Add(txtRawOCRLog);
                y += 210;

                // ── Mapping Log ───────────────────────────────────────────────
                UIHelper.CreateSectionLabel(
                    pnlOCR,
                    "✅ Chi tiết quét OCR (Mapping kết quả):",
                    ref y
                );
                y -= 15;

                var mapSearchPanel = CreateSearchBarForMap(pnlOCR, y);
                y += 32;

                txtProcessLog = new RichTextBox
                {
                    Location = new Point(10, y),
                    Width = this.ClientSize.Width - 50,
                    Height = 400,
                    ReadOnly = true,
                    BackColor = Color.White,
                    Font = new Font("Courier New", 8),
                    BorderStyle = BorderStyle.FixedSingle,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                };
                pnlOCR.Controls.Add(txtProcessLog);
                y += 410;

                // ── Export mode ───────────────────────────────────────────────
                UIHelper.CreateSectionLabel(
                    pnlOCR,
                    "📤 Chọn cách ghi Sheet Excel khi xuất:",
                    ref y
                );
                y -= 15;

                var grpExportMode = new Panel
                {
                    Location = new Point(10, y),
                    Width = this.ClientSize.Width > 50 ? this.ClientSize.Width - 50 : 500,
                    Height = 100,
                    BackColor = SystemColors.Control,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                };
                pnlOCR.Controls.Add(grpExportMode);

                var rdoByInvoiceDate = new RadioButton
                {
                    Text = "Theo ngày hóa đơn  (VD: 26 → sheet 26-02, 27 → sheet 27-02 ...)",
                    Location = new Point(0, 4),
                    Height = 22,
                    Font = new Font("Arial", 9),
                    Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                    AutoSize = true,
                };
                var rdoToday = new RadioButton
                {
                    Text =
                        "Ngày hôm nay  (tất cả vào sheet " + DateTime.Now.ToString("dd-MM") + ")",
                    Location = new Point(0, 32),
                    Height = 22,
                    Font = new Font("Arial", 9),
                    Checked = true,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                    AutoSize = true,
                };
                var rdoCustomDate = new RadioButton
                {
                    Text = "Ngày khác (VD: 25-02):",
                    Location = new Point(0, 60),
                    Height = 22,
                    Width = 170,
                    Font = new Font("Arial", 9),
                    AutoSize = false,
                };
                var txtCustomDate = new TextBox
                {
                    Text = "",
                    PlaceholderText = "dd-MM",
                    Location = new Point(174, 59),
                    Width = 120,
                    Height = 22,
                    Font = new Font("Arial", 9),
                    Enabled = false,
                };

                _exportUseToday = true; // sync với rdoToday.Checked = true ở trên

                rdoByInvoiceDate.CheckedChanged += (s, e) =>
                {
                    if (rdoByInvoiceDate.Checked)
                    {
                        _exportUseToday = false;
                        txtCustomDate.Enabled = false;
                    }
                };
                rdoToday.CheckedChanged += (s, e) =>
                {
                    if (rdoToday.Checked)
                    {
                        _exportUseToday = true;
                        txtCustomDate.Enabled = false;
                    }
                };
                rdoCustomDate.CheckedChanged += (s, e) =>
                {
                    if (rdoCustomDate.Checked)
                    {
                        _exportUseToday = null;
                        txtCustomDate.Enabled = true;
                        txtCustomDate.Focus();
                    }
                };
                txtCustomDate.TextChanged += (s, e) =>
                {
                    _exportCustomDate = txtCustomDate.Text.Trim();
                };

                grpExportMode.Controls.Add(rdoByInvoiceDate);
                grpExportMode.Controls.Add(rdoToday);
                grpExportMode.Controls.Add(rdoCustomDate);
                grpExportMode.Controls.Add(txtCustomDate);
                y += 102;

                // ── Export button ─────────────────────────────────────────────
                var btnExportOCR = UIHelper.CreateButton(
                    "💾 XUẤT EXCEL",
                    Color.LightGreen,
                    10,
                    y,
                    150,
                    35
                );
                btnExportOCR.Click += (s, e) => ExportMappedDataToExcel();
                pnlOCR.Controls.Add(btnExportOCR);

                // ── Tag refs ──────────────────────────────────────────────────
                pnlOCR.Tag = new Dictionary<string, object>
                {
                    { "rawLog", txtRawOCRLog },
                    { "mappingLog", txtProcessLog },
                };

                // ── Responsive resize ─────────────────────────────────────────
                pnlOCR.Resize += (s, e) =>
                {
                    if (txtRawOCRLog != null)
                        txtRawOCRLog.Width = pnlOCR.ClientSize.Width - 30;
                    if (txtProcessLog != null)
                        txtProcessLog.Width = pnlOCR.ClientSize.Width - 30;
                    if (rawSearchPanel != null)
                        rawSearchPanel.Width = pnlOCR.ClientSize.Width - 20;
                    if (mapSearchPanel != null)
                        mapSearchPanel.Width = pnlOCR.ClientSize.Width - 20;
                    if (grpExportMode != null)
                        grpExportMode.Width = pnlOCR.ClientSize.Width - 20;
                };

                tabOCR.Controls.Clear();
                tabOCR.Controls.Add(pnlOCR);

                Debug.WriteLine("✅ OCR Tab UI initialized");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ Error initializing OCR Tab UI: {ex.Message}\n{ex.StackTrace}");
                // Show lỗi trên tab thay vì để trắng
                var lblErr = new Label
                {
                    Text = $"❌ Lỗi khởi tạo tab OCR:\n{ex.Message}",
                    Dock = DockStyle.Fill,
                    ForeColor = Color.Red,
                    Font = new Font("Arial", 10),
                    TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                    Padding = new Padding(10),
                };
                tabOCR.Controls.Clear();
                tabOCR.Controls.Add(lblErr);
            }
        }

        // ─── Search bar helpers ────────────────────────────────────────────────
        // (delegate sang UIHelper — giữ ở đây vì gắn liền với txtRawOCRLog / txtProcessLog)
        private Panel CreateSearchBarForRaw(Panel parent, int y) =>
            UIHelper.CreateRichTextBoxSearchBar(parent, y, () => txtRawOCRLog);

        private Panel CreateSearchBarForMap(Panel parent, int y) =>
            UIHelper.CreateRichTextBoxSearchBar(parent, y, () => txtProcessLog);
    }
}

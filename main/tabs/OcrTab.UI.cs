using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;

namespace TextInputter
{
    /// <summary>
    /// OcrTab UI — control field declarations + InitializeOCRTab().
    /// Logic (ProcessImages, SelectOCRFolder, ExportMappedDataToExcel...) ở OcrTab.cs.
    /// </summary>
    public partial class MainForm
    {
        // ─── Controls thuộc OCR Tab ────────────────────────────────────────────
        private TextBox txtNguoiDiOCR;
        private TextBox txtNguoiLayOCR;
        private RichTextBox txtRawOCRLog;
        private RichTextBox txtProcessLog;

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
                    Dock       = DockStyle.Fill,
                    AutoScroll = true,
                    BackColor  = SystemColors.Control,
                    Padding    = new Padding(10)
                };

                int y = 10;

                // ── Title ──────────────────────────────────────────────────────
                UIHelper.CreateSectionLabel(pnlOCR, "🔍 OCR Processing", ref y);
                y -= 15;

                // ── Folder selection ───────────────────────────────────────────
                Label lblFolderInfo = new Label
                {
                    Text     = "Chọn folder ảnh để quét OCR tự động",
                    AutoSize = true,
                    Location = new Point(10, y),
                    Font     = new Font("Arial", 10, FontStyle.Bold)
                };
                pnlOCR.Controls.Add(lblFolderInfo);
                y += 25;

                var btnSelectFolder = UIHelper.CreateButton("📂 Chọn Folder", Color.LightBlue, 10, y, 130, 35);
                btnSelectFolder.Click += (s, e) => SelectOCRFolder();
                pnlOCR.Controls.Add(btnSelectFolder);

                var btnStartScan = UIHelper.CreateButton("▶ Bắt Đầu Quét", Color.LightGreen, 150, y, 130, 35);
                btnStartScan.Click += (s, e) => btnStart_Click(null, EventArgs.Empty);
                pnlOCR.Controls.Add(btnStartScan);
                y += 45;

                // ── Người Đi ──────────────────────────────────────────────────
                UIHelper.CreateSectionLabel(pnlOCR, "Thông tin NGƯỜI ĐI & NGƯỜI LẤY (bắt buộc):", ref y);
                y -= 15;

                pnlOCR.Controls.Add(new Label
                {
                    Text     = "Người Đi:",
                    AutoSize = true,
                    Location = new Point(10, y),
                    Font     = new Font("Arial", 9, FontStyle.Bold)
                });

                txtNguoiDiOCR = new TextBox
                {
                    Location    = new Point(10, y + 20),
                    Width       = pnlOCR.ClientSize.Width - 20,
                    Height      = 28,
                    BorderStyle = BorderStyle.FixedSingle,
                    Font        = new Font("Arial", 11)
                };
                pnlOCR.Controls.Add(txtNguoiDiOCR);
                y += 60;

                // ── Người Lấy ─────────────────────────────────────────────────
                pnlOCR.Controls.Add(new Label
                {
                    Text     = "Người Lấy:",
                    AutoSize = true,
                    Location = new Point(10, y),
                    Font     = new Font("Arial", 9, FontStyle.Bold)
                });

                txtNguoiLayOCR = new TextBox
                {
                    Location    = new Point(10, y + 20),
                    Width       = pnlOCR.ClientSize.Width - 20,
                    Height      = 28,
                    BorderStyle = BorderStyle.FixedSingle,
                    Font        = new Font("Arial", 11)
                };
                pnlOCR.Controls.Add(txtNguoiLayOCR);
                y += 60;

                // ── Raw OCR Log ───────────────────────────────────────────────
                UIHelper.CreateSectionLabel(pnlOCR, "📋 Raw OCR Text (Kết quả OCR thô):", ref y);
                y -= 15;

                var rawSearchPanel = CreateSearchBarForRaw(pnlOCR, y);
                y += 32;

                txtRawOCRLog = new RichTextBox
                {
                    Location    = new Point(10, y),
                    Width       = pnlOCR.ClientSize.Width - 30,
                    Height      = 200,
                    ReadOnly    = true,
                    BackColor   = Color.White,
                    Font        = new Font("Courier New", 8),
                    BorderStyle = BorderStyle.FixedSingle
                };
                pnlOCR.Controls.Add(txtRawOCRLog);
                y += 210;

                // ── Mapping Log ───────────────────────────────────────────────
                UIHelper.CreateSectionLabel(pnlOCR, "✅ Chi tiết quét OCR (Mapping kết quả):", ref y);
                y -= 15;

                var mapSearchPanel = CreateSearchBarForMap(pnlOCR, y);
                y += 32;

                txtProcessLog = new RichTextBox
                {
                    Location    = new Point(10, y),
                    Width       = pnlOCR.ClientSize.Width - 30,
                    Height      = 400,
                    ReadOnly    = true,
                    BackColor   = Color.White,
                    Font        = new Font("Courier New", 8),
                    BorderStyle = BorderStyle.FixedSingle
                };
                pnlOCR.Controls.Add(txtProcessLog);
                y += 410;

                // ── Export button ─────────────────────────────────────────────
                var btnExportOCR = UIHelper.CreateButton("💾 XUẤT EXCEL", Color.LightGreen, 10, y, 150, 35);
                btnExportOCR.Click += (s, e) => ExportMappedDataToExcel();
                pnlOCR.Controls.Add(btnExportOCR);

                // ── Tag refs ──────────────────────────────────────────────────
                pnlOCR.Tag = new Dictionary<string, object>
                {
                    { "rawLog",     txtRawOCRLog },
                    { "mappingLog", txtProcessLog }
                };

                // ── Responsive resize ─────────────────────────────────────────
                pnlOCR.Resize += (s, e) =>
                {
                    if (txtNguoiDiOCR  != null) txtNguoiDiOCR.Width  = pnlOCR.ClientSize.Width - 20;
                    if (txtNguoiLayOCR != null) txtNguoiLayOCR.Width = pnlOCR.ClientSize.Width - 20;
                    if (txtRawOCRLog   != null) txtRawOCRLog.Width   = pnlOCR.ClientSize.Width - 30;
                    if (txtProcessLog  != null) txtProcessLog.Width  = pnlOCR.ClientSize.Width - 30;
                    if (rawSearchPanel != null) rawSearchPanel.Width = pnlOCR.ClientSize.Width - 20;
                    if (mapSearchPanel != null) mapSearchPanel.Width = pnlOCR.ClientSize.Width - 20;
                };

                tabOCR.Controls.Clear();
                tabOCR.Controls.Add(pnlOCR);

                Debug.WriteLine("✅ OCR Tab UI initialized");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ Error initializing OCR Tab UI: {ex.Message}");
            }
        }

        // ─── Search bar helpers ────────────────────────────────────────────────
        // (delegate sang UIHelper — giữ ở đây vì gắn liền với txtRawOCRLog / txtProcessLog)
        private Panel CreateSearchBarForRaw(Panel parent, int y)
            => UIHelper.CreateRichTextBoxSearchBar(parent, y, () => txtRawOCRLog);

        private Panel CreateSearchBarForMap(Panel parent, int y)
            => UIHelper.CreateRichTextBoxSearchBar(parent, y, () => txtProcessLog);
    }
}

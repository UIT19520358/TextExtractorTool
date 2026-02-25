using System;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;

namespace TextInputter
{
    /// <summary>
    /// ManualInputTab UI — InitializeManualInputTab() + CreateMandatoryField() helper.
    /// Logic (SaveManualEntry) ở ManualInputTab.cs.
    /// </summary>
    public partial class MainForm
    {
        // ─── Init ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Khởi tạo tab nhập thủ công với 17 trường bắt buộc (highlight vàng).
        /// Gọi từ MainForm constructor sau InitializeComponent().
        /// </summary>
        private void InitializeManualInputTab()
        {
            try
            {
                Panel pnlManualInput = new Panel
                {
                    Dock       = DockStyle.Fill,
                    AutoScroll = true,
                    BackColor  = SystemColors.Control,
                    Padding    = new Padding(10)
                };

                int y = 10;

                UIHelper.CreateSectionLabel(pnlManualInput, "✋ Nhập Dữ Liệu Thủ Công (17 Trường Bắt Buộc)", ref y);
                y -= 15;

                pnlManualInput.Controls.Add(new Label
                {
                    Text      = "⭐ Tất cả các trường màu vàng là bắt buộc phải điền",
                    AutoSize  = true,
                    ForeColor = Color.OrangeRed,
                    Font      = new Font("Arial", 9, FontStyle.Bold),
                    Location  = new Point(10, y)
                });
                y += 25;

                // ── Section 1: Basic Info ──────────────────────────────────────
                UIHelper.CreateSectionLabel(pnlManualInput, "📋 Thông Tin Cơ Bản:", ref y);
                y -= 15;

                var txtTinhTrang = CreateMandatoryField(pnlManualInput, "[1] Tình Trạng TT:", ref y);
                var txtThuTu     = CreateMandatoryField(pnlManualInput, "[2] Thứ:", ref y);
                var txtNgay      = CreateMandatoryField(pnlManualInput, "[3] Ngày (DD-MM-YYYY):", ref y);
                var txtMa        = CreateMandatoryField(pnlManualInput, "[4] Mã:", ref y);

                // ── Section 2: Address ─────────────────────────────────────────
                UIHelper.CreateSectionLabel(pnlManualInput, "📍 Địa Chỉ:", ref y);
                y -= 15;

                var txtSoNha    = CreateMandatoryField(pnlManualInput, "[5] Số Nhà:", ref y);
                var txtTenDuong = CreateMandatoryField(pnlManualInput, "[6] Tên Đường:", ref y);
                var txtQuan     = CreateMandatoryField(pnlManualInput, "[7] Quận:", ref y);

                // ── Section 3: Money ───────────────────────────────────────────
                UIHelper.CreateSectionLabel(pnlManualInput, "💰 Tiền Tệ:", ref y);
                y -= 15;

                var txtTienThu  = CreateMandatoryField(pnlManualInput, "[8] Tiền Thu:", ref y);
                var txtTienShip = CreateMandatoryField(pnlManualInput, "[9] Tiền Ship:", ref y);
                var txtTienHang = CreateMandatoryField(pnlManualInput, "[10] Tiền Hàng:", ref y);

                // ── Section 4: People & Status ─────────────────────────────────
                UIHelper.CreateSectionLabel(pnlManualInput, "👥 Người Liên Quan & Trạng Thái:", ref y);
                y -= 15;

                var txtNguoiDi  = CreateMandatoryField(pnlManualInput, "[11] Người Đi:", ref y);
                var txtNguoiLay = CreateMandatoryField(pnlManualInput, "[12] Người Lấy:", ref y);
                var txtGhiChu   = CreateMandatoryField(pnlManualInput, "[13] Ghi Chú:", ref y);
                var txtUng      = CreateMandatoryField(pnlManualInput, "[14] Ứng tiền:", ref y);
                var txtHang     = CreateMandatoryField(pnlManualInput, "[15] Hàng tồn:", ref y);
                var txtFail     = CreateMandatoryField(pnlManualInput, "[16] Fail:", ref y);
                var txtNote     = CreateMandatoryField(pnlManualInput, "[17] Ghi Chú Thêm:", ref y);

                // ── Buttons ────────────────────────────────────────────────────
                y += 10;

                var btnSaveManual = UIHelper.CreateButton("💾 Lưu", Color.LightGreen, 10, y, 100, 35);
                btnSaveManual.Click += (s, e) => SaveManualEntry(
                    txtTinhTrang.Text, txtThuTu.Text, txtNgay.Text, txtMa.Text,
                    txtSoNha.Text, txtTenDuong.Text, txtQuan.Text,
                    txtTienThu.Text, txtTienShip.Text, txtTienHang.Text,
                    txtNguoiDi.Text, txtNguoiLay.Text, txtGhiChu.Text,
                    txtUng.Text, txtHang.Text, txtFail.Text, txtNote.Text);
                pnlManualInput.Controls.Add(btnSaveManual);

                var btnClearManual = UIHelper.CreateButton("🔄 Xóa", Color.LightCoral, 120, y, 100, 35);
                btnClearManual.Click += (s, e) =>
                {
                    foreach (var txt in new[] { txtTinhTrang, txtThuTu, txtNgay, txtMa,
                                                txtSoNha, txtTenDuong, txtQuan,
                                                txtTienThu, txtTienShip, txtTienHang,
                                                txtNguoiDi, txtNguoiLay, txtGhiChu,
                                                txtUng, txtHang, txtFail, txtNote })
                        txt.Clear();
                };
                pnlManualInput.Controls.Add(btnClearManual);

                tabManualInput.Controls.Clear();
                tabManualInput.Controls.Add(pnlManualInput);

                Debug.WriteLine("✅ Manual Input Tab UI initialized (17 fields)");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ Error initializing Manual Input Tab UI: {ex.Message}");
            }
        }

        /// <summary>
        /// Tạo một field bắt buộc: Label + TextBox highlight vàng.
        /// </summary>
        private TextBox CreateMandatoryField(Panel panel, string labelText, ref int yPos, bool isMultiline = false)
        {
            panel.Controls.Add(new Label
            {
                Text      = labelText,
                AutoSize  = true,
                Location  = new Point(10, yPos),
                Font      = new Font("Arial", 9, FontStyle.Bold),
                ForeColor = Color.Black
            });
            yPos += 20;

            var txt = new TextBox
            {
                Location    = new Point(10, yPos),
                Width       = panel.ClientSize.Width - 30,
                Height      = isMultiline ? 60 : 25,
                Multiline   = isMultiline,
                BackColor   = Color.Yellow,
                Font        = new Font("Arial", 9),
                BorderStyle = BorderStyle.FixedSingle
            };
            panel.Controls.Add(txt);
            yPos += isMultiline ? 70 : 35;

            return txt;
        }
    }
}

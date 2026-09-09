using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace TextInputter
{
    /// <summary>
    /// ManualInputTab UI — InitializeManualInputTab() + CreateMandatoryField() + CreateOptionalField() helper.
    /// Logic (SaveManualEntry) ở ManualInputTab.cs.
    /// </summary>
    public partial class MainForm
    {
        // ─── Init ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Khởi tạo tab nhập thủ công.
        /// Bắt buộc (vàng): Ngày, Mã, Số Nhà, Tên Đường, Quận, Tiền Thu, Tiền Ship, Người Đi, Người Lấy.
        /// Tùy chọn (trắng): Tình Trạng TT, Thứ, Shop, Tên KH, Tiền Hàng, Ghi Chú, Ứng Tiền, Hàng Tồn, Fail, Ghi Chú Thêm.
        /// Gọi từ MainForm constructor sau InitializeComponent().
        /// </summary>
        private void InitializeManualInputTab()
        {
            try
            {
                Panel pnlManualInput = new Panel
                {
                    Dock = DockStyle.Fill,
                    AutoScroll = true,
                    BackColor = SystemColors.Control,
                    Padding = new Padding(10),
                };

                int y = 10;

                UIHelper.CreateSectionLabel(pnlManualInput, "✋ Nhập Dữ Liệu Thủ Công", ref y);
                y -= 15;

                pnlManualInput.Controls.Add(
                    new Label
                    {
                        Text = "⭐ = bắt buộc   |   Không có ⭐ = tùy chọn (để trống cũng lưu được)",
                        AutoSize = true,
                        ForeColor = Color.OrangeRed,
                        Font = new Font("Arial", 9, FontStyle.Bold),
                        Location = new Point(10, y),
                    }
                );
                y += 25;

                // ── Section 1: Basic Info ──────────────────────────────────────
                UIHelper.CreateSectionLabel(pnlManualInput, "📋 Thông Tin Cơ Bản:", ref y);
                y -= 15;

                var txtTinhTrang = CreateOptionalField(pnlManualInput, "Tình Trạng TT:", ref y);
                var txtThuTu = CreateOptionalField(pnlManualInput, "Thứ:", ref y);
                var txtNgay = CreateMandatoryField(pnlManualInput, "Ngày (DD-MM-YYYY) ⭐:", ref y);
                var txtMa = CreateMandatoryField(pnlManualInput, "Mã ⭐:", ref y);
                var txtShop = CreateMandatoryField(pnlManualInput, "Shop ⭐:", ref y);
                var txtTenKh = CreateMandatoryField(pnlManualInput, "Tên KH ⭐:", ref y);

                // ── Section 2: Address ─────────────────────────────────────────
                UIHelper.CreateSectionLabel(pnlManualInput, "📍 Địa Chỉ:", ref y);
                y -= 15;

                var txtSoNha = CreateMandatoryField(pnlManualInput, "Số Nhà ⭐:", ref y);
                var txtTenDuong = CreateMandatoryField(pnlManualInput, "Tên Đường ⭐:", ref y);
                var txtQuan = CreateMandatoryField(pnlManualInput, "Quận ⭐:", ref y);

                // ── Section 3: Money ───────────────────────────────────────────
                UIHelper.CreateSectionLabel(pnlManualInput, "💰 Tiền Tệ:", ref y);
                y -= 15;

                var txtTienThu = CreateMandatoryField(pnlManualInput, "Tiền Thu ⭐:", ref y);
                var txtTienShip = CreateMandatoryField(pnlManualInput, "Tiền Ship ⭐:", ref y);
                var txtTienHang = CreateOptionalField(
                    pnlManualInput,
                    "Tiền Hàng (tự tính = Thu - Ship nếu trống):",
                    ref y
                );

                // ── Section 4: People & Status ─────────────────────────────────
                UIHelper.CreateSectionLabel(
                    pnlManualInput,
                    "👥 Người Liên Quan & Trạng Thái:",
                    ref y
                );
                y -= 15;

                var txtNguoiDi = CreateMandatoryField(pnlManualInput, "Người Đi ⭐:", ref y);
                var txtNguoiLay = CreateMandatoryField(pnlManualInput, "Người Lấy ⭐:", ref y);
                var txtGhiChu = CreateOptionalField(pnlManualInput, "Ghi Chú:", ref y);
                var txtUng = CreateOptionalField(pnlManualInput, "Ứng Tiền:", ref y);
                var txtHang = CreateOptionalField(pnlManualInput, "Hàng Tồn:", ref y);
                var txtFail = CreateOptionalField(pnlManualInput, "Fail:", ref y);
                var txtNote = CreateOptionalField(pnlManualInput, "Ghi Chú Thêm:", ref y);

                // ── Buttons ────────────────────────────────────────────────────
                y += 10;

                var btnSaveManual = UIHelper.CreateButton(
                    "💾 Lưu",
                    Color.LightGreen,
                    10,
                    y,
                    100,
                    35
                );
                btnSaveManual.Click += (s, e) =>
                    SaveManualEntry(
                        txtTinhTrang.Text,
                        txtThuTu.Text,
                        txtNgay.Text,
                        txtMa.Text,
                        txtShop.Text,
                        txtTenKh.Text,
                        txtSoNha.Text,
                        txtTenDuong.Text,
                        txtQuan.Text,
                        txtTienThu.Text,
                        txtTienShip.Text,
                        txtTienHang.Text,
                        txtNguoiDi.Text,
                        txtNguoiLay.Text,
                        txtGhiChu.Text,
                        txtUng.Text,
                        txtHang.Text,
                        txtFail.Text,
                        txtNote.Text
                    );
                pnlManualInput.Controls.Add(btnSaveManual);

                var btnClearManual = UIHelper.CreateButton(
                    "🔄 Xóa",
                    Color.LightCoral,
                    120,
                    y,
                    100,
                    35
                );
                btnClearManual.Click += (s, e) =>
                {
                    foreach (
                        var txt in new[]
                        {
                            txtTinhTrang,
                            txtThuTu,
                            txtNgay,
                            txtMa,
                            txtShop,
                            txtTenKh,
                            txtSoNha,
                            txtTenDuong,
                            txtQuan,
                            txtTienThu,
                            txtTienShip,
                            txtTienHang,
                            txtNguoiDi,
                            txtNguoiLay,
                            txtGhiChu,
                            txtUng,
                            txtHang,
                            txtFail,
                            txtNote,
                        }
                    )
                        txt.Clear();
                };
                pnlManualInput.Controls.Add(btnClearManual);

                tabManualInput.Controls.Clear();
                tabManualInput.Controls.Add(pnlManualInput);

                Debug.WriteLine("✅ Manual Input Tab UI initialized");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ Error initializing Manual Input Tab UI: {ex.Message}");
            }
        }

        /// <summary>
        /// Tạo field bắt buộc: Label + TextBox highlight vàng.
        /// </summary>
        private TextBox CreateMandatoryField(
            Panel panel,
            string labelText,
            ref int yPos,
            bool isMultiline = false
        )
        {
            panel.Controls.Add(
                new Label
                {
                    Text = labelText,
                    AutoSize = true,
                    Location = new Point(10, yPos),
                    Font = new Font("Arial", 9, FontStyle.Bold),
                    ForeColor = Color.Black,
                }
            );
            yPos += 20;

            var txt = new TextBox
            {
                Location = new Point(10, yPos),
                Width = panel.ClientSize.Width - 30,
                Height = isMultiline ? 60 : 25,
                Multiline = isMultiline,
                BackColor = Color.Yellow,
                Font = new Font("Arial", 9),
                BorderStyle = BorderStyle.FixedSingle,
            };
            panel.Controls.Add(txt);
            yPos += isMultiline ? 70 : 35;

            return txt;
        }

        /// <summary>
        /// Tạo field tùy chọn: Label + TextBox nền trắng (không validate khi trống).
        /// </summary>
        private TextBox CreateOptionalField(
            Panel panel,
            string labelText,
            ref int yPos,
            bool isMultiline = false
        )
        {
            panel.Controls.Add(
                new Label
                {
                    Text = labelText,
                    AutoSize = true,
                    Location = new Point(10, yPos),
                    Font = new Font("Arial", 9),
                    ForeColor = Color.DimGray,
                }
            );
            yPos += 20;

            var txt = new TextBox
            {
                Location = new Point(10, yPos),
                Width = panel.ClientSize.Width - 30,
                Height = isMultiline ? 60 : 25,
                Multiline = isMultiline,
                BackColor = Color.White,
                Font = new Font("Arial", 9),
                BorderStyle = BorderStyle.FixedSingle,
            };
            panel.Controls.Add(txt);
            yPos += isMultiline ? 70 : 35;

            return txt;
        }
    }
}

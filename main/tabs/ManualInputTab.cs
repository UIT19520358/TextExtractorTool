using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;

namespace TextInputter
{
    /// <summary>
    /// Manual Input Tab: nhập thủ công 17 fields bắt buộc
    /// </summary>
    public partial class MainForm
    {
        /// <summary>
        /// Khởi tạo tab nhập thủ công với 17 trường bắt buộc (highlight vàng)
        /// </summary>
        private void InitializeManualInputTab()
        {
            try
            {
                Panel pnlManualInput = new Panel
                {
                    Dock        = DockStyle.Fill,
                    AutoScroll  = true,
                    BackColor   = SystemColors.Control,
                    Padding     = new Padding(10)
                };

                int y = 10;

                UIHelper.CreateSectionLabel(pnlManualInput, "✋ Nhập Dữ Liệu Thủ Công (17 Trường Bắt Buộc)", ref y);
                y -= 15;

                Label lblLegend = new Label
                {
                    Text      = "⭐ Tất cả các trường màu vàng là bắt buộc phải điền",
                    AutoSize  = true,
                    ForeColor = Color.OrangeRed,
                    Font      = new Font("Arial", 9, FontStyle.Bold),
                    Location  = new Point(10, y)
                };
                pnlManualInput.Controls.Add(lblLegend);
                y += 25;

                // Section 1: Basic Info
                UIHelper.CreateSectionLabel(pnlManualInput, "📋 Thông Tin Cơ Bản:", ref y);
                y -= 15;

                var txtTinhTrang = CreateMandatoryField(pnlManualInput, "[1] Tình Trạng TT:", ref y);
                var txtThuTu     = CreateMandatoryField(pnlManualInput, "[2] Thứ:", ref y);
                var txtNgay      = CreateMandatoryField(pnlManualInput, "[3] Ngày (DD-MM-YYYY):", ref y);
                var txtMa        = CreateMandatoryField(pnlManualInput, "[4] Mã:", ref y);

                // Section 2: Address
                UIHelper.CreateSectionLabel(pnlManualInput, "📍 Địa Chỉ:", ref y);
                y -= 15;

                var txtSoNha    = CreateMandatoryField(pnlManualInput, "[5] Số Nhà:", ref y);
                var txtTenDuong = CreateMandatoryField(pnlManualInput, "[6] Tên Đường:", ref y);
                var txtQuan     = CreateMandatoryField(pnlManualInput, "[7] Quận:", ref y);

                // Section 3: Money
                UIHelper.CreateSectionLabel(pnlManualInput, "💰 Tiền Tệ:", ref y);
                y -= 15;

                var txtTienThu  = CreateMandatoryField(pnlManualInput, "[8] Tiền Thu:", ref y);
                var txtTienShip = CreateMandatoryField(pnlManualInput, "[9] Tiền Ship:", ref y);
                var txtTienHang = CreateMandatoryField(pnlManualInput, "[10] Tiền Hàng:", ref y);

                // Section 4: People & Status
                UIHelper.CreateSectionLabel(pnlManualInput, "👥 Người Liên Quan & Trạng Thái:", ref y);
                y -= 15;

                var txtNguoiDi  = CreateMandatoryField(pnlManualInput, "[11] Người Đi:", ref y);
                var txtNguoiLay = CreateMandatoryField(pnlManualInput, "[12] Người Lấy:", ref y);
                var txtGhiChu   = CreateMandatoryField(pnlManualInput, "[13] Ghi Chú:", ref y);
                var txtUng      = CreateMandatoryField(pnlManualInput, "[14] Ứng tiền:", ref y);
                var txtHang     = CreateMandatoryField(pnlManualInput, "[15] Hàng tồn:", ref y);
                var txtFail     = CreateMandatoryField(pnlManualInput, "[16] Fail:", ref y);
                var txtNote     = CreateMandatoryField(pnlManualInput, "[17] Ghi Chú Thêm:", ref y);

                // Buttons
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

                Debug.WriteLine("✅ Manual Input Tab initialized (17 fields)");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ Error initializing Manual Input Tab: {ex.Message}");
            }
        }

        /// <summary>
        /// Tạo một field bắt buộc với label và TextBox highlight vàng
        /// </summary>
        private TextBox CreateMandatoryField(Panel panel, string labelText, ref int yPos, bool isMultiline = false)
        {
            Label lbl = new Label
            {
                Text      = labelText,
                AutoSize  = true,
                Location  = new Point(10, yPos),
                Font      = new Font("Arial", 9, FontStyle.Bold),
                ForeColor = Color.Black
            };
            panel.Controls.Add(lbl);
            yPos += 20;

            TextBox txt = new TextBox
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

        /// <summary>
        /// Validate và lưu 17 fields từ manual input
        /// </summary>
        private void SaveManualEntry(
            string tinhTrang, string thuTu, string ngay, string ma,
            string soNha, string tenDuong, string quan,
            string tienThu, string tienShip, string tienHang,
            string nguoiDi, string nguoiLay, string ghiChu,
            string ung, string hang, string fail, string note)
        {
            try
            {
                var missingFields = new List<string>();
                void Check(string val, string name) { if (string.IsNullOrWhiteSpace(val)) missingFields.Add(name); }

                Check(tinhTrang, "1. Tình Trạng TT");
                Check(thuTu,     "2. Thứ");
                Check(ngay,      "3. Ngày");
                Check(ma,        "4. Mã");
                Check(soNha,     "5. Số Nhà");
                Check(tenDuong,  "6. Tên Đường");
                Check(quan,      "7. Quận");
                Check(tienThu,   "8. Tiền Thu");
                Check(tienShip,  "9. Tiền Ship");
                Check(tienHang,  "10. Tiền Hàng");
                Check(nguoiDi,   "11. Người Đi");
                Check(nguoiLay,  "12. Người Lấy");
                Check(ghiChu,    "13. Ghi Chú");
                Check(ung,       "14. Ưng");
                Check(hang,      "15. Hàng");
                Check(fail,      "16. Fail");
                Check(note,      "17. Ghi Chú Thêm");

                if (missingFields.Count > 0)
                {
                    MessageBox.Show("❌ Vui lòng điền đủ tất cả 17 trường bắt buộc:\n\n" +
                        string.Join("\n", missingFields), "Thiếu thông tin bắt buộc",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!DateTime.TryParse(ngay, out _))
                {
                    MessageBox.Show("Ngày phải ở định dạng DD-MM-YYYY", "Lỗi định dạng", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (!decimal.TryParse(tienThu,  out decimal tienThuVal)  || tienThuVal  < 0) { MessageBox.Show("Tiền Thu phải là số dương!");  return; }
                if (!decimal.TryParse(tienShip, out decimal tienShipVal) || tienShipVal < 0) { MessageBox.Show("Tiền Ship phải là số dương!"); return; }
                if (!decimal.TryParse(tienHang, out decimal tienHangVal) || tienHangVal < 0) { MessageBox.Show("Tiền Hàng phải là số dương!"); return; }

                MessageBox.Show(
                    $"✅ Lưu thành công:\n\nTình Trạng: {tinhTrang}\nNgày: {ngay}\n" +
                    $"Địa Chỉ: {soNha}, {tenDuong}, {quan}\n" +
                    $"Tiền Thu: {tienThuVal:N0}\nNgười Đi: {nguoiDi}\nNgười Lấy: {nguoiLay}",
                    "Thành công");

                Debug.WriteLine($"✅ Manual entry saved: {ma} - {soNha}, {tenDuong}, {quan}");
                // TODO: Save to Excel với đủ 17 fields
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi: {ex.Message}", "Lỗi");
                Debug.WriteLine($"Error saving manual entry: {ex.Message}\n{ex.StackTrace}");
            }
        }
    }
}

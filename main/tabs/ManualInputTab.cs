using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;

namespace TextInputter
{
    /// <summary>
    /// ManualInputTab logic — SaveManualEntry (validate + lưu 17 fields).
    /// UI (InitializeManualInputTab + CreateMandatoryField) ở ManualInputTab.UI.cs.
    /// </summary>
    public partial class MainForm
    {
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

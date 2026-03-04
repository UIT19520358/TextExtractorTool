using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace TextInputter
{
    /// <summary>
    /// ManualInputTab logic — SaveManualEntry (validate + ghi Excel).
    /// UI (InitializeManualInputTab + CreateMandatoryField) ở ManualInputTab.UI.cs.
    /// </summary>
    public partial class MainForm
    {
        /// <summary>
        /// Validate và lưu fields từ manual input vào file Excel (giống Xuất Excel của OCR tab).
        /// </summary>
        private void SaveManualEntry(
            string tinhTrang,
            string thuTu,
            string ngay,
            string ma,
            string shop,
            string tenKh,
            string soNha,
            string tenDuong,
            string quan,
            string tienThu,
            string tienShip,
            string tienHang,
            string nguoiDi,
            string nguoiLay,
            string ghiChu,
            string ung,
            string hang,
            string fail,
            string note
        )
        {
            try
            {
                // ── Validate — chỉ các field bắt buộc ────────────────────────
                var missingFields = new List<string>();
                void Check(string val, string name)
                {
                    if (string.IsNullOrWhiteSpace(val))
                        missingFields.Add(name);
                }

                // Bắt buộc (giống OCR tab)
                Check(ngay, "Ngày");
                Check(ma, "Mã");
                Check(shop, "Shop");
                Check(tenKh, "Tên KH");
                Check(soNha, "Số Nhà");
                Check(tenDuong, "Tên Đường");
                Check(quan, "Quận");
                Check(tienThu, "Tiền Thu");
                Check(tienShip, "Tiền Ship");
                Check(nguoiDi, "Người Đi");
                Check(nguoiLay, "Người Lấy");

                // Tùy chọn: tinhTrang, thuTu, tienHang, ghiChu, ung, hang, fail, note
                // → không validate, để trống vẫn lưu được

                if (missingFields.Count > 0)
                {
                    MessageBox.Show(
                        "❌ Vui lòng điền đủ các trường bắt buộc (⭐):\n\n"
                            + string.Join("\n", missingFields),
                        "Thiếu thông tin bắt buộc",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }

                if (
                    !DateTime.TryParseExact(
                        ngay,
                        new[] { "dd-MM-yyyy", "d-M-yyyy", "dd/MM/yyyy" },
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None,
                        out _
                    )
                )
                {
                    MessageBox.Show(
                        "Ngày phải ở định dạng DD-MM-YYYY",
                        "Lỗi định dạng",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }
                if (!decimal.TryParse(tienThu, out decimal tienThuVal) || tienThuVal < 0)
                {
                    MessageBox.Show("Tiền Thu phải là số dương!");
                    return;
                }
                if (!decimal.TryParse(tienShip, out decimal tienShipVal) || tienShipVal < 0)
                {
                    MessageBox.Show("Tiền Ship phải là số dương!");
                    return;
                }

                // Tiền Hàng: dùng giá trị nhập nếu có, không thì tự tính = Thu - Ship
                decimal tienHangVal;
                if (
                    string.IsNullOrWhiteSpace(tienHang)
                    || !decimal.TryParse(tienHang, out tienHangVal)
                    || tienHangVal < 0
                )
                    tienHangVal = tienThuVal - tienShipVal;

                // ── Chọn file Excel ───────────────────────────────────────────
                using var openDialog = new OpenFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                    Title = "Chọn file Excel để lưu dữ liệu",
                    InitialDirectory = Path.Combine(
                        Directory.GetCurrentDirectory(),
                        "data",
                        "sample",
                        "excel"
                    ),
                };
                if (openDialog.ShowDialog() != DialogResult.OK)
                    return;

                string excelPath = openDialog.FileName;

                // ── Xác định tên sheet từ ngày nhập ──────────────────────────
                // ngay đã được validate format DD-MM-YYYY ở trên
                var ngayParts = ngay.Split('-');
                string sheetName = $"{ngayParts[0]}-{ngayParts[1]}"; // VD: "11-02"

                DateTime.TryParseExact(
                    sheetName,
                    "dd-MM",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    out DateTime sheetDate
                );

                // ── Ghi vào Excel ─────────────────────────────────────────────
                // ⚠️ HARDCODED: 20-column header — phải khớp với template Excel của khách
                var headers = new[]
                {
                    "Tình trạng TT",
                    "SHOP",
                    "TÊN KH",
                    "MÃ",
                    "SỐ NHÀ",
                    "TÊN ĐƯỜNG",
                    "QUẬN",
                    "TIỀN THU",
                    "TIỀN SHIP",
                    "TIỀN HÀNG",
                    "NGƯỜI ĐI",
                    "NGƯỜI LẤY",
                    "NGÀY LẤY",
                    "GHI CHÚ",
                    "ỨNG TIỀN",
                    "HÀNG TỒN",
                    "FAIL",
                    "Column1",
                    "Column2",
                    "Column3",
                };

                using var workbook = new XLWorkbook(excelPath);
                bool isNewSheet = !workbook.TryGetWorksheet(sheetName, out var worksheet);
                if (isNewSheet)
                {
                    worksheet = workbook.Worksheets.Add(sheetName);
                    // Header row
                    for (int col = 0; col < headers.Length; col++)
                    {
                        var cell = worksheet.Cell(1, col + 1);
                        cell.Value = headers[col];
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                    }
                    // Row 2: THU x / NGAY x-x label
                    string thuText =
                        sheetDate.DayOfWeek == DayOfWeek.Sunday
                            ? "CHU NHAT"
                            : "THU " + ((int)sheetDate.DayOfWeek + 1);
                    worksheet.Cell(2, 2).Value = thuText;
                    worksheet.Cell(2, 2).Style.Font.Bold = true;
                    worksheet.Cell(2, 3).Value = $"NGAY {sheetDate.Day}-{sheetDate.Month}";
                    worksheet.Cell(2, 3).Style.Font.Bold = true;
                }

                // Data bắt đầu từ row 3; tìm row cuối để append
                int currentRow = 3;
                var lastUsed = worksheet.LastRowUsed();
                if (lastUsed != null && lastUsed.RowNumber() >= 3)
                    currentRow = lastUsed.RowNumber() + 1;

                // Upsert theo MÃ
                int targetRow = -1;
                foreach (var row in worksheet.RowsUsed())
                {
                    if (row.RowNumber() <= 2)
                        continue;
                    if (row.Cell(4).GetString() == ma)
                    {
                        targetRow = row.RowNumber();
                        break;
                    }
                }
                bool isUpdate = targetRow > 0;
                if (!isUpdate)
                    targetRow = currentRow;

                worksheet.Cell(targetRow, 1).Value = tinhTrang;
                worksheet.Cell(targetRow, 2).Value = shop;
                worksheet.Cell(targetRow, 3).Value = tenKh;
                worksheet.Cell(targetRow, 4).Value = ma;
                worksheet.Cell(targetRow, 5).Value = soNha;
                worksheet.Cell(targetRow, 6).Value = tenDuong;
                worksheet.Cell(targetRow, 7).Value = quan;
                worksheet.Cell(targetRow, 8).Value = tienThuVal;
                worksheet.Cell(targetRow, 9).Value = tienShipVal;
                worksheet.Cell(targetRow, 10).Value = tienHangVal;
                worksheet.Cell(targetRow, 11).Value = nguoiDi;
                worksheet.Cell(targetRow, 12).Value = nguoiLay;
                worksheet.Cell(targetRow, 13).Value = ngay;
                worksheet.Cell(targetRow, 14).Value = ghiChu;
                worksheet.Cell(targetRow, 15).Value = ung;
                worksheet.Cell(targetRow, 16).Value = hang;
                worksheet.Cell(targetRow, 17).Value = fail;
                worksheet.Cell(targetRow, 18).Value = note;

                workbook.SaveAs(excelPath);

                string action = isUpdate ? "✏️ Ghi đè" : "➕ Thêm mới";
                MessageBox.Show(
                    $"✅ Lưu thành công!\n\n{action}: {ma}\n📅 Sheet: {sheetName}\n📂 File: {Path.GetFileName(excelPath)}",
                    "✅ Thành công",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                Debug.WriteLine(
                    $"✅ Manual entry saved: {ma} → sheet '{sheetName}' row {targetRow} ({(isUpdate ? "update" : "insert")})"
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌ Lỗi: {ex.Message}",
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                Debug.WriteLine($"Error saving manual entry: {ex.Message}\n{ex.StackTrace}");
            }
        }
    }
}

using System.Collections.Generic;

namespace TextInputter
{
    /// <summary>
    /// Invoice Tab — shared state.
    /// Logic được chia thành các file partial:
    ///   InvoiceTab.ExcelViewer.cs  — mở / đọc / lưu Excel viewer
    ///   InvoiceTab.Calculate.cs    — tính tiền từ Excel → daily report
    ///   InvoiceTab.Report.cs       — hiển thị + lưu daily report
    /// </summary>
    public partial class MainForm
    {
        // ─── Shared data class ────────────────────────────────────────────────

        private class DailyReportData
        {
            public string Date { get; set; }
            public decimal TongTienThu { get; set; }   // Tổng tiền thu (cột TIỀN THU)
            public decimal TongTienShip { get; set; }  // Tổng tiền ship (cột TIỀN SHIP)
            public decimal KhoanTruShip { get; set; }  // -(TongShip - SoDon×5), số âm
            public decimal TongKetCuoi { get; set; }   // TongHangDuong + row âm
            public decimal SoDon { get; set; }

            // Các row âm (đơn trả, đơn cũ ck...) lấy từ Excel
            public System.Collections.Generic.List<(string Label, decimal Amount)> NegativeRows { get; set; } = new();

            // Report nhỏ theo từng người đi: Key = tên người, Value = (TienThu, TienShip, SoDon)
            public System.Collections.Generic.Dictionary<
                string,
                (decimal TienThu, decimal TienShip, decimal SoDon)
            > ReportByNguoiDi { get; set; } =
                new System.Collections.Generic.Dictionary<string, (decimal, decimal, decimal)>(
                    System.StringComparer.OrdinalIgnoreCase
                );
        }

        private DailyReportData currentDailyReport;
    }
}

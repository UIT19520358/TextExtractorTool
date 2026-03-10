using System.Collections.Generic;

namespace TextInputter
{
    /// <summary>
    /// Invoice Tab — shared state.
    /// Logic được chia thành các file partial:
    ///   InvoiceTab.ExcelViewer.cs  — mở / đọc / lưu Excel viewer
    ///   InvoiceTab.Calculate.cs    — tính tiền từ Excel → daily report
    ///   InvoiceTab.Report.cs       — hiển thị + lưu daily report
    ///   InvoiceTab.Returns.cs      — đánh dấu hàng trả + import đối soát
    /// </summary>
    public partial class MainForm
    {
        // ─── Per-person detailed report ───────────────────────────────────────

        private class NguoiDiDetail
        {
            public decimal TienThu { get; set; }       // Tổng tiền thu
            public decimal TienShip { get; set; }      // Tổng tiền ship
            public decimal SoDon { get; set; }         // Tổng số đơn
            public int SoDonGop { get; set; }          // Số đơn gộp (giao 1 lần cho nhiều đơn cùng địa chỉ)
            public int SoDonTra { get; set; }          // Số đơn trả (đã đánh dấu FAIL=xx)
            public decimal TienShipTru { get; set; }   // -(TongShip - SoDonGiao × 5k), số âm
            public decimal TienLay { get; set; }       // -((SoDon - SoDonTra - SoDonGop) × 2k), số âm
            public decimal TienDonTra { get; set; }    // Tổng tiền trừ đơn trả, số âm
            public bool IsAnTam { get; set; }          // true = An Tâm → skip ship/lấy/trả calculation

            /// <summary>Số đơn giao thực tế = SoDon − SoDonGop</summary>
            public decimal SoDonGiao => SoDon - SoDonGop;

            /// <summary>Chi tiết từng đơn trả: (MÃ HĐ, TiềnThu, ShipFee theo quận, Tiền trừ)</summary>
            public System.Collections.Generic.List<(string Ma, decimal TienThu, decimal ShipFee, decimal Deduction)>
                DonTraDetails { get; set; } = new();
        }

        // ─── Shared data class ────────────────────────────────────────────────

        private class DailyReportData
        {
            public string Date { get; set; }
            public decimal TongTienThu { get; set; }   // Tổng tiền thu (cột TIỀN THU)
            public decimal TongTienShip { get; set; }  // Tổng tiền ship (cột TIỀN SHIP)
            public decimal KhoanTruShip { get; set; }  // -(TongShip - SoDon×5), số âm
            public decimal TongKetCuoi { get; set; }   // TongHangDuong + row âm
            public decimal SoDon { get; set; }
            public int TotalDonGop { get; set; }       // Tổng đơn gộp (toàn bộ)
            public int TotalDonTra { get; set; }       // Tổng đơn trả (toàn bộ)

            // LEFT "Đơn trả": -SUMIFS(TIỀN HÀNG, FAIL="xx", ỨNG TIỀN="x") — matching Excel formula
            public decimal TongTienHangDonTra { get; set; }
            public int SoDonTraLeft { get; set; }      // COUNTIFS(FAIL="xx", ỨNG TIỀN="x")

            // Hàng tồn (carry-over từ ngày trước, HÀNG TỒN=x) — loại khỏi LEFT summary
            public decimal TongTienThuHangTon { get; set; }
            public decimal TongTienShipHangTon { get; set; }
            public int SoDonHangTon { get; set; }

            // Các row âm (đơn trả, đơn cũ ck...) lấy từ Excel
            public System.Collections.Generic.List<(string Label, decimal Amount)> NegativeRows { get; set; } = new();

            // Report chi tiết theo từng người đi (thay thế tuple cũ)
            public System.Collections.Generic.Dictionary<string, NguoiDiDetail> DetailByNguoiDi { get; set; } =
                new System.Collections.Generic.Dictionary<string, NguoiDiDetail>(
                    System.StringComparer.OrdinalIgnoreCase
                );
        }

        private DailyReportData currentDailyReport;
    }
}

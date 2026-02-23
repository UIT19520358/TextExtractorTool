using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace TextInputter.Services
{
    /// <summary>
    /// Xử lý business logic cho tính toán
    /// </summary>
    public class BusinessLogicService
    {
        /// <summary>
        /// Tính tổng tiền từ data
        /// </summary>
        public decimal CalculateTotalAmount(DataGridView dgv)
        {
            decimal total = 0;

            try
            {
                for (int row = 0; row < dgv.Rows.Count; row++)
                {
                    var amountCell = dgv.Rows[row].Cells["TIỀN HÀNG"];
                    if (amountCell != null && decimal.TryParse(amountCell.Value?.ToString(), out decimal amount))
                    {
                        total += amount;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi tính toán:\n{ex.Message}", "Lỗi");
            }

            return total;
        }

        /// <summary>
        /// Tính tổng số đơn
        /// </summary>
        public decimal CalculateTotalOrders(DataGridView dgv)
        {
            decimal total = 0;

            try
            {
                for (int row = 0; row < dgv.Rows.Count; row++)
                {
                    var orderCell = dgv.Rows[row].Cells["SỐ ĐƠN"];
                    if (orderCell != null && decimal.TryParse(orderCell.Value?.ToString(), out decimal orders))
                    {
                        total += orders;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi tính toán:\n{ex.Message}", "Lỗi");
            }

            return total;
        }

        /// <summary>
        /// Tính adjustment (điều chỉnh, khấu hao, v.v.)
        /// </summary>
        public Dictionary<string, decimal> CalculateAdjustments(DataGridView dgv)
        {
            var adjustments = new Dictionary<string, decimal>();

            try
            {
                for (int row = 0; row < dgv.Rows.Count; row++)
                {
                    var shopCell = dgv.Rows[row].Cells["SHOP"];
                    var amountCell = dgv.Rows[row].Cells["TIỀN HÀNG"];

                    // Nếu SHOP trống = đây là hàng adjustment
                    if (shopCell != null && string.IsNullOrWhiteSpace(shopCell.Value?.ToString()))
                    {
                        var descCell = dgv.Rows[row].Cells["Tình trạng"];
                        string adjustmentType = descCell?.Value?.ToString()?.Trim() ?? "Other";

                        if (decimal.TryParse(amountCell?.Value?.ToString(), out decimal amount))
                        {
                            if (!adjustments.ContainsKey(adjustmentType))
                                adjustments[adjustmentType] = 0;

                            adjustments[adjustmentType] += amount;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi tính adjustment:\n{ex.Message}", "Lỗi");
            }

            return adjustments;
        }

        /// <summary>
        /// Lọc data rows (exclude adjustments)
        /// </summary>
        public List<DataGridViewRow> GetDataRows(DataGridView dgv)
        {
            var dataRows = new List<DataGridViewRow>();

            try
            {
                for (int row = 0; row < dgv.Rows.Count; row++)
                {
                    var shopCell = dgv.Rows[row].Cells["SHOP"];
                    
                    // Data rows phải có SHOP (không trống)
                    if (shopCell != null && !string.IsNullOrWhiteSpace(shopCell.Value?.ToString()))
                    {
                        dataRows.Add(dgv.Rows[row]);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi lọc data:\n{ex.Message}", "Lỗi");
            }

            return dataRows;
        }

        /// <summary>
        /// Lấy adjustment rows
        /// </summary>
        public List<DataGridViewRow> GetAdjustmentRows(DataGridView dgv)
        {
            var adjustmentRows = new List<DataGridViewRow>();

            try
            {
                for (int row = 0; row < dgv.Rows.Count; row++)
                {
                    var shopCell = dgv.Rows[row].Cells["SHOP"];
                    
                    // Adjustment rows có SHOP trống
                    if (shopCell != null && string.IsNullOrWhiteSpace(shopCell.Value?.ToString()))
                    {
                        adjustmentRows.Add(dgv.Rows[row]);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi lấy adjustment rows:\n{ex.Message}", "Lỗi");
            }

            return adjustmentRows;
        }

        /// <summary>
        /// Validate data trước khi tính toán
        /// </summary>
        public bool ValidateData(DataGridView dgv)
        {
            if (dgv.Rows.Count == 0)
            {
                MessageBox.Show("❌ Không có dữ liệu để tính toán!", "Lỗi");
                return false;
            }

            try
            {
                // Check required columns
                var requiredColumns = new[] { "SHOP", "TIỀN HÀNG", "SỐ ĐƠN" };
                foreach (var column in requiredColumns)
                {
                    if (!dgv.Columns.Contains(column))
                    {
                        MessageBox.Show($"❌ Cột '{column}' không tìm thấy!", "Lỗi");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi validate:\n{ex.Message}", "Lỗi");
                return false;
            }

            return true;
        }
    }
}

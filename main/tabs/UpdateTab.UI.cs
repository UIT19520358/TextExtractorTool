using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using TextInputter.Services;

namespace TextInputter
{
    /// <summary>
    /// UpdateTab UI + Logic — cho phép cập nhật hàng loạt theo MÃ HĐ.
    /// UI: ô nhập mã (multi-line), checkboxes từng cột, textbox input tương ứng.
    /// Logic: UpdateInvoiceFields() trong ExcelInvoiceService.
    /// </summary>
    public partial class MainForm
    {
        // ── State ─────────────────────────────────────────────────────────────
        // Map: tên cột → (CheckBox, TextBox input)
        private readonly Dictionary<string, (CheckBox chk, TextBox txt)> _updateFieldControls =
            new Dictionary<string, (CheckBox, TextBox)>(StringComparer.OrdinalIgnoreCase);

        // ── Init ──────────────────────────────────────────────────────────────
        private void InitializeUpdateTab()
        {
            try
            {
                var pnl = new Panel
                {
                    Dock = DockStyle.Fill,
                    AutoScroll = true,
                    BackColor = SystemColors.Control,
                    Padding = new Padding(12),
                };

                int y = 10;

                // ── Tiêu đề ──────────────────────────────────────────────────
                pnl.Controls.Add(
                    new Label
                    {
                        Text = "✏️ Cập Nhật Hàng Loạt Theo Mã HĐ",
                        Font = new Font("Arial", 12, FontStyle.Bold),
                        AutoSize = true,
                        Location = new Point(10, y),
                    }
                );
                y += 32;

                pnl.Controls.Add(
                    new Label
                    {
                        Text =
                            "Nhập một hoặc nhiều mã HĐ (mỗi mã một dòng), chọn cột cần sửa rồi bấm Cập Nhật.",
                        Font = new Font("Arial", 9),
                        ForeColor = Color.DimGray,
                        AutoSize = true,
                        Location = new Point(10, y),
                    }
                );
                y += 22;

                // ── Ô nhập MÃ ────────────────────────────────────────────────
                pnl.Controls.Add(
                    new Label
                    {
                        Text = "📋 Mã HĐ ⭐ (mỗi mã một dòng):",
                        Font = new Font("Arial", 9, FontStyle.Bold),
                        AutoSize = true,
                        Location = new Point(10, y),
                    }
                );
                y += 20;

                var txtMaList = new TextBox
                {
                    Location = new Point(10, y),
                    Width = 400,
                    Height = 90,
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    Font = new Font("Consolas", 9),
                    BackColor = Color.LightYellow,
                    BorderStyle = BorderStyle.FixedSingle,
                    PlaceholderText = "HD133277\nHD133304\nHD133306",
                };
                pnl.Controls.Add(txtMaList);
                y += 100;

                // ── Divider ───────────────────────────────────────────────────
                pnl.Controls.Add(
                    new Label
                    {
                        Text = "──────────────────────────────────────────────────────────────",
                        Font = new Font("Arial", 8),
                        ForeColor = Color.LightGray,
                        AutoSize = true,
                        Location = new Point(10, y),
                    }
                );
                y += 18;

                pnl.Controls.Add(
                    new Label
                    {
                        Text = "✅ Chọn cột cần chỉnh sửa:",
                        Font = new Font("Arial", 9, FontStyle.Bold),
                        AutoSize = true,
                        Location = new Point(10, y),
                    }
                );
                y += 22;

                // ── Checkboxes + TextBoxes cho từng cột ──────────────────────
                _updateFieldControls.Clear();
                int colW = 280; // width mỗi cột checkbox+input
                int cols = 2; // số cột layout (2 cột song song)
                int col = 0;

                foreach (var fieldName in ExcelInvoiceService.EditableColumnNames)
                {
                    int x = 10 + col * (colW + 20);

                    var chk = new CheckBox
                    {
                        Text = fieldName,
                        Font = new Font("Arial", 9, FontStyle.Bold),
                        AutoSize = true,
                        Location = new Point(x, y),
                        ForeColor = Color.DarkSlateBlue,
                    };

                    var txt = new TextBox
                    {
                        Location = new Point(x + 20, y + 22),
                        Width = colW - 30,
                        Height = 24,
                        Font = new Font("Arial", 9),
                        BackColor = Color.White,
                        Visible = false, // ẩn mặc định
                        BorderStyle = BorderStyle.FixedSingle,
                        PlaceholderText = $"Giá trị mới cho {fieldName}...",
                    };

                    // Toggle textbox: ẩn/hiện khi check/uncheck
                    chk.CheckedChanged += (s, e) =>
                    {
                        txt.Visible = chk.Checked;
                        if (!chk.Checked)
                            txt.Clear();
                    };

                    pnl.Controls.Add(chk);
                    pnl.Controls.Add(txt);
                    _updateFieldControls[fieldName] = (chk, txt);

                    col++;
                    if (col >= cols)
                    {
                        col = 0;
                        y += 58; // xuống hàng mới
                    }
                }
                if (col != 0)
                    y += 58; // flush hàng cuối chưa đầy

                // ── Nút chọn tất cả / bỏ chọn ────────────────────────────────
                y += 6;
                var btnSelectAll = new Button
                {
                    Text = "✓ Chọn tất cả",
                    Location = new Point(10, y),
                    Size = new Size(120, 28),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.LightSteelBlue,
                    Font = new Font("Arial", 8, FontStyle.Bold),
                    Cursor = Cursors.Hand,
                };
                btnSelectAll.Click += (s, e) =>
                {
                    foreach (var (chk, _) in _updateFieldControls.Values)
                        chk.Checked = true;
                };
                pnl.Controls.Add(btnSelectAll);

                var btnDeselectAll = new Button
                {
                    Text = "✗ Bỏ tất cả",
                    Location = new Point(140, y),
                    Size = new Size(110, 28),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.LightGray,
                    Font = new Font("Arial", 8),
                    Cursor = Cursors.Hand,
                };
                btnDeselectAll.Click += (s, e) =>
                {
                    foreach (var (chk, _) in _updateFieldControls.Values)
                        chk.Checked = false;
                };
                pnl.Controls.Add(btnDeselectAll);
                y += 38;

                // ── Divider ───────────────────────────────────────────────────
                pnl.Controls.Add(
                    new Label
                    {
                        Text = "──────────────────────────────────────────────────────────────",
                        Font = new Font("Arial", 8),
                        ForeColor = Color.LightGray,
                        AutoSize = true,
                        Location = new Point(10, y),
                    }
                );
                y += 18;

                // ── Nút Cập Nhật ─────────────────────────────────────────────
                var btnUpdate = new Button
                {
                    Text = "✏️ Cập Nhật",
                    Location = new Point(10, y),
                    Size = new Size(130, 36),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.MediumSeaGreen,
                    ForeColor = Color.White,
                    Font = new Font("Arial", 10, FontStyle.Bold),
                    Cursor = Cursors.Hand,
                };
                btnUpdate.FlatAppearance.BorderSize = 0;
                btnUpdate.Click += (s, e) => ExecuteUpdate(txtMaList.Text);
                pnl.Controls.Add(btnUpdate);

                var btnClearUpdate = new Button
                {
                    Text = "🔄 Xóa form",
                    Location = new Point(150, y),
                    Size = new Size(110, 36),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.LightCoral,
                    ForeColor = Color.White,
                    Font = new Font("Arial", 9),
                    Cursor = Cursors.Hand,
                };
                btnClearUpdate.FlatAppearance.BorderSize = 0;
                btnClearUpdate.Click += (s, e) =>
                {
                    txtMaList.Clear();
                    foreach (var (chk, txt) in _updateFieldControls.Values)
                    {
                        chk.Checked = false;
                        txt.Clear();
                    }
                };
                pnl.Controls.Add(btnClearUpdate);
                y += 46;

                // ── Label kết quả ─────────────────────────────────────────────
                var lblUpdateResult = new Label
                {
                    Name = "lblUpdateResult",
                    Text = "",
                    AutoSize = false,
                    Location = new Point(10, y),
                    Size = new Size(700, 40),
                    Font = new Font("Arial", 9, FontStyle.Bold),
                    ForeColor = Color.DarkGreen,
                };
                pnl.Controls.Add(lblUpdateResult);

                tabUpdate.Controls.Clear();
                tabUpdate.Controls.Add(pnl);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ UpdateTab init error: {ex.Message}");
            }
        }

        // ── Logic ─────────────────────────────────────────────────────────────
        private void ExecuteUpdate(string maRawText)
        {
            // 1. Lấy label kết quả
            Label lblResult = null;
            foreach (Control c in tabUpdate.Controls)
                if (c is Panel p)
                    foreach (Control pc in p.Controls)
                        if (pc is Label lb && lb.Name == "lblUpdateResult")
                            lblResult = lb;

            void SetResult(string msg, Color color)
            {
                if (lblResult != null)
                {
                    lblResult.Text = msg;
                    lblResult.ForeColor = color;
                }
            }

            // 2. Validate: cần file Excel đang mở
            if (string.IsNullOrWhiteSpace(currentExcelFilePath))
            {
                MessageBox.Show(
                    "⚠️ Chưa mở file Excel.\nHãy mở file Excel ở tab Excel Viewer trước.",
                    "Chưa có file",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            // 2.5. Chọn sheet để cập nhật
            string selectedSheet = null;
            try
            {
                using var wb = new ClosedXML.Excel.XLWorkbook(currentExcelFilePath);
                var sheets = wb.Worksheets.Select(ws => ws.Name).ToList();
                
                if (sheets.Count == 0)
                {
                    MessageBox.Show("File Excel không có sheet nào!", "Lỗi",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                if (sheets.Count == 1)
                {
                    selectedSheet = sheets[0];
                }
                else
                {
                    // Hỏi user chọn sheet
                    using var dlg = new Form
                    {
                        Text = "Chọn sheet để cập nhật",
                        Size = new System.Drawing.Size(360, 180),
                        StartPosition = FormStartPosition.CenterParent,
                        FormBorderStyle = FormBorderStyle.FixedDialog,
                        MaximizeBox = false, MinimizeBox = false,
                    };
                    var lbl = new Label
                    {
                        Text = "Chọn sheet:",
                        Left = 12, Top = 12, AutoSize = true,
                        Font = new Font("Arial", 9, FontStyle.Bold),
                    };
                    var cmb = new ComboBox
                    {
                        Left = 12, Top = 34, Width = 320,
                        DropDownStyle = ComboBoxStyle.DropDownList,
                    };
                    cmb.Items.AddRange(sheets.Cast<object>().ToArray());
                    cmb.SelectedIndex = 0;
                    
                    var btnOk = new Button
                    {
                        Text = "OK", Left = 180, Top = 80, Width = 75,
                        DialogResult = DialogResult.OK,
                    };
                    var btnCancel = new Button
                    {
                        Text = "Hủy", Left = 265, Top = 80, Width = 75,
                        DialogResult = DialogResult.Cancel,
                    };
                    
                    dlg.Controls.AddRange(new Control[] { lbl, cmb, btnOk, btnCancel });
                    dlg.AcceptButton = btnOk;
                    dlg.CancelButton = btnCancel;
                    
                    if (dlg.ShowDialog() != DialogResult.OK) return;
                    selectedSheet = cmb.SelectedItem?.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi đọc file Excel: {ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrEmpty(selectedSheet))
            {
                MessageBox.Show("Chưa chọn sheet!", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 3. Parse danh sách MÃ
            var maList = maRawText
                .Split(new[] { '\n', '\r', ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(m => m.Trim())
                .Where(m => !string.IsNullOrEmpty(m))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (maList.Count == 0)
            {
                MessageBox.Show(
                    "⚠️ Vui lòng nhập ít nhất một Mã HĐ.",
                    "Thiếu Mã HĐ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            // 4. Thu thập các field được check + có giá trị
            var fieldValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in _updateFieldControls)
            {
                var (chk, txt) = kv.Value;
                if (chk.Checked)
                    fieldValues[kv.Key] = txt.Text; // cho phép ghi chuỗi rỗng (xoá giá trị)
            }

            if (fieldValues.Count == 0)
            {
                MessageBox.Show(
                    "⚠️ Chưa chọn cột nào để cập nhật.\nHãy tick vào ít nhất một cột.",
                    "Chưa chọn cột",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            // 5. Xác nhận
            string preview = string.Join(", ", fieldValues.Keys);
            string confirm =
                $"Sẽ cập nhật {maList.Count} đơn hàng:\n\n"
                + $"  Mã: {string.Join(", ", maList.Take(5))}{(maList.Count > 5 ? $" ... (+{maList.Count - 5} nữa)" : "")}\n\n"
                + $"  Cột: {preview}\n\n"
                + "Tiếp tục?";

            if (
                MessageBox.Show(
                    confirm,
                    "Xác nhận cập nhật",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                ) != DialogResult.Yes
            )
                return;

            // 6. Gọi service
            try
            {
                var svc = new ExcelInvoiceService(currentExcelFilePath);
                var (updated, notFound) = svc.UpdateInvoiceFields(selectedSheet, maList, fieldValues);

                string msg = $"✅ Đã cập nhật {updated}/{maList.Count} đơn (sheet: {selectedSheet}).";
                if (notFound.Count > 0)
                    msg += $"\n⚠️ Không tìm thấy: {string.Join(", ", notFound)}";

                SetResult(msg, notFound.Count == 0 ? Color.DarkGreen : Color.DarkOrange);

                if (notFound.Count > 0)
                    MessageBox.Show(
                        $"✅ Cập nhật thành công: {updated} đơn.\n\n"
                            + $"⚠️ Không tìm thấy {notFound.Count} mã:\n{string.Join("\n", notFound)}",
                        "Kết quả cập nhật",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                else
                    MessageBox.Show(
                        $"✅ Đã cập nhật {updated} đơn thành công!",
                        "Thành công",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );

                // Reload Excel viewer nếu đang mở cùng file
                ReloadExcelIfOpen();
            }
            catch (Exception ex)
            {
                SetResult($"❌ Lỗi: {ex.Message}", Color.Red);
                MessageBox.Show(
                    $"❌ Lỗi khi cập nhật:\n{ex.Message}",
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                System.Diagnostics.Debug.WriteLine(
                    $"UpdateTab error: {ex.Message}\n{ex.StackTrace}"
                );
            }
        }

        /// <summary>Reload lại Excel Viewer nếu file hiện tại đang được mở.</summary>
        private void ReloadExcelIfOpen()
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(currentExcelFilePath))
                    LoadExcelFile(currentExcelFilePath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ReloadExcel error: {ex.Message}");
            }
        }
    }
}

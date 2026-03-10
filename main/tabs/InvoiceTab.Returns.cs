using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace TextInputter
{
    // ─── Return Marking + Đối Soát Import ───────────────────────────────────────
    public partial class MainForm
    {
        /// <summary>
        /// Hiển thị dialog đánh dấu đơn trả.
        /// User nhập MÃ HĐ → app tìm trong sheet hiện tại → đánh dấu X/xx.
        /// Nếu không tìm thấy → cho import từ file đối soát.
        /// </summary>
        private void ShowReturnDialog()
        {
            using var dlg = new Form
            {
                Text = "📋 Đánh dấu đơn trả / đơn CK",
                Size = new Size(500, 400),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.White,
            };

            // ── Loại đánh dấu ─────────────────────────────────────────────────
            var lblType = new Label
            {
                Text = "Loại:",
                Left = 12, Top = 14, Width = 40,
                Font = new Font("Arial", 10),
            };
            var rbTraTrongNgay = new RadioButton
            {
                Text = "Trả trong ngày (X ứng tiền, xx fail)",
                Left = 55, Top = 12, Width = 300,
                Checked = true,
                Font = new Font("Arial", 9),
            };
            var rbTraNgayTruoc = new RadioButton
            {
                Text = "Trả ngày trước / CK (X ứng tiền + hàng tồn, xx fail)",
                Left = 55, Top = 34, Width = 400,
                Font = new Font("Arial", 9),
            };

            // ── Nhập MÃ HĐ ───────────────────────────────────────────────────
            var lblMa = new Label
            {
                Text = "MÃ HĐ (nhập nhiều dòng, mỗi dòng 1 mã):",
                Left = 12, Top = 62, AutoSize = true,
                Font = new Font("Arial", 10),
            };
            var txtMa = new TextBox
            {
                Left = 12, Top = 84, Width = 460, Height = 120,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 10),
            };

            // ── Kết quả ───────────────────────────────────────────────────────
            var lblResult = new Label
            {
                Text = "",
                Left = 12, Top = 210, Width = 460, Height = 80,
                Font = new Font("Arial", 9),
                ForeColor = Color.DarkGreen,
            };

            // ── Buttons ───────────────────────────────────────────────────────
            var btnMark = new Button
            {
                Text = "✅ Đánh dấu",
                Left = 12, Top = 310, Width = 120, Height = 35,
                BackColor = Color.FromArgb(40, 40, 40),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Arial", 10),
            };
            btnMark.FlatAppearance.BorderSize = 0;

            var btnImport = new Button
            {
                Text = "📥 Import đối soát",
                Left = 140, Top = 310, Width = 160, Height = 35,
                BackColor = Color.FromArgb(40, 40, 40),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Arial", 10),
            };
            btnImport.FlatAppearance.BorderSize = 0;

            var btnClose = new Button
            {
                Text = "Đóng",
                Left = 400, Top = 310, Width = 75, Height = 35,
                DialogResult = DialogResult.Cancel,
            };

            // ── Event handlers ────────────────────────────────────────────────
            btnMark.Click += (s, e) =>
            {
                var maList = txtMa.Text
                    .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(m => m.Trim())
                    .Where(m => !string.IsNullOrEmpty(m))
                    .ToList();

                if (maList.Count == 0)
                {
                    MessageBox.Show("Vui lòng nhập ít nhất 1 MÃ HĐ!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                bool includeHangTon = rbTraNgayTruoc.Checked;
                var (found, notFound) = MarkReturnsInSourceGrid(maList, includeHangTon);

                string msg = $"✅ Đã đánh dấu: {found.Count} đơn";
                if (notFound.Count > 0)
                    msg += $"\n❌ Không tìm thấy: {notFound.Count} đơn\n   → {string.Join(", ", notFound)}";
                lblResult.Text = msg;
                lblResult.ForeColor = notFound.Count > 0 ? Color.DarkRed : Color.DarkGreen;
            };

            btnImport.Click += (s, e) =>
            {
                // Truyền danh sách MÃ HĐ đã nhập (nếu có) để auto-search trong đối soát
                var maList = txtMa.Text
                    .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(m => m.Trim())
                    .Where(m => !string.IsNullOrEmpty(m))
                    .ToList();
                ImportFromDoiSoat(lblResult, maList);
            };

            dlg.Controls.AddRange(new Control[]
            {
                lblType, rbTraTrongNgay, rbTraNgayTruoc,
                lblMa, txtMa, lblResult,
                btnMark, btnImport, btnClose,
            });
            dlg.CancelButton = btnClose;
            dlg.ShowDialog();
        }

        /// <summary>
        /// Tìm và đánh dấu đơn trả trong source DataGridView (tabExcelSheets).
        /// Trả về (danh sách MÃ tìm thấy, danh sách MÃ không tìm thấy).
        /// </summary>
        private (List<string> found, List<string> notFound) MarkReturnsInSourceGrid(
            List<string> maList, bool includeHangTon)
        {
            var found = new List<string>();
            var notFound = new List<string>();

            // Lấy source grid hiện tại
            DataGridView sourceGrid = GetCurrentSourceGrid();
            if (sourceGrid == null)
            {
                notFound.AddRange(maList);
                return (found, notFound);
            }

            // Detect relevant columns
            int colMa = -1, colUngTien = -1, colHangTon = -1, colFail = -1;
            for (int c = 0; c < sourceGrid.Columns.Count; c++)
            {
                string h = sourceGrid.Columns[c].HeaderText.ToLower();
                if (h.Contains("mã"))
                    colMa = c;
                if (h.Contains("ứng tiền") || h.Contains("ung tien"))
                    colUngTien = c;
                if (h.Contains("hàng tồn") || h.Contains("hang ton"))
                    colHangTon = c;
                if (h.Contains("fail"))
                    colFail = c;
            }

            if (colMa < 0)
            {
                MessageBox.Show("Không tìm thấy cột MÃ trong sheet hiện tại!",
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                notFound.AddRange(maList);
                return (found, notFound);
            }

            foreach (string ma in maList)
            {
                bool matched = false;
                for (int i = 0; i < sourceGrid.Rows.Count; i++)
                {
                    var row = sourceGrid.Rows[i];
                    if (row.IsNewRow)
                        continue;
                    string cellMa = colMa < row.Cells.Count
                        ? (row.Cells[colMa].Value?.ToString() ?? "").Trim()
                        : "";
                    if (!cellMa.Equals(ma, StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Đánh dấu
                    if (colUngTien >= 0 && colUngTien < row.Cells.Count)
                        row.Cells[colUngTien].Value = "x";
                    if (includeHangTon && colHangTon >= 0 && colHangTon < row.Cells.Count)
                        row.Cells[colHangTon].Value = "x";
                    if (colFail >= 0 && colFail < row.Cells.Count)
                        row.Cells[colFail].Value = "xx";

                    // Tô nền đỏ nhạt để dễ nhận biết
                    for (int c = 0; c < row.Cells.Count; c++)
                        row.Cells[c].Style.BackColor = Color.FromArgb(255, 220, 220);

                    matched = true;
                    found.Add(ma);
                    break;
                }
                if (!matched)
                    notFound.Add(ma);
            }

            return (found, notFound);
        }

        /// <summary>
        /// Import đơn từ file đối soát (Excel) vào source grid hiện tại.
        /// Flow: mở file → quét TẤT CẢ sheets → auto-match MÃ HĐ → user confirm → copy + đánh dấu.
        /// </summary>
        /// <param name="lblResult">Label hiển thị kết quả.</param>
        /// <param name="searchMaList">Danh sách MÃ HĐ cần tìm (nếu rỗng → hiện toàn bộ).</param>
        private void ImportFromDoiSoat(Label lblResult, List<string> searchMaList = null)
        {
            DataGridView sourceGrid = GetCurrentSourceGrid();
            if (sourceGrid == null)
            {
                MessageBox.Show("Chưa mở file Excel nào!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Mở file đối soát
            string doiSoatPath;
            using (var ofd = new OpenFileDialog
            {
                Title = "Chọn file đối soát",
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
            })
            {
                if (ofd.ShowDialog() != DialogResult.OK)
                    return;
                doiSoatPath = ofd.FileName;
            }

            try
            {
                var allRows = new List<Dictionary<string, string>>();
                List<string> doiSoatHeaders = null;

                using (var workbook = new XLWorkbook(doiSoatPath))
                {
                    // Quét TẤT CẢ sheets (không chỉ sheet cuối)
                    foreach (var ws in workbook.Worksheets)
                    {
                        var usedRange = ws.RangeUsed();
                        if (usedRange == null)
                            continue;

                        // Detect header row
                        int headerRow = 1;
                        int colCount = usedRange.ColumnCount();
                        int rowCount = usedRange.RowCount();
                        for (int r = 1; r <= Math.Min(5, rowCount); r++)
                        {
                            bool found = false;
                            for (int c = 1; c <= Math.Min(colCount, 10); c++)
                            {
                                string v = ws.Cell(r, c).GetString().Trim();
                                if (v.Contains("MÃ", StringComparison.OrdinalIgnoreCase)
                                    || v.Contains("SHOP", StringComparison.OrdinalIgnoreCase)
                                    || v.Contains("TÊN", StringComparison.OrdinalIgnoreCase))
                                {
                                    headerRow = r;
                                    found = true;
                                    break;
                                }
                            }
                            if (found) break;
                        }

                        // Dùng headers từ sheet đầu tiên có data (tất cả sheets cùng format)
                        if (doiSoatHeaders == null)
                        {
                            doiSoatHeaders = new List<string>();
                            for (int c = 1; c <= colCount; c++)
                                doiSoatHeaders.Add(ws.Cell(headerRow, c).GetString().Trim());
                        }

                        // Detect MÃ column index để filter
                        int maColIdx = -1;
                        for (int c = 1; c <= colCount; c++)
                        {
                            string h = ws.Cell(headerRow, c).GetString().Trim().ToUpper();
                            if (h == "MÃ" || h == "MÃ HĐ" || h == "MA")
                            {
                                maColIdx = c;
                                break;
                            }
                        }

                        // Read data rows
                        for (int r = headerRow + 1; r <= rowCount; r++)
                        {
                            var rowData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            bool hasData = false;
                            for (int c = 0; c < doiSoatHeaders.Count; c++)
                            {
                                string val = ws.Cell(r, c + 1).GetString().Trim();
                                if (!string.IsNullOrEmpty(val))
                                    hasData = true;
                                rowData[doiSoatHeaders[c]] = val;
                            }

                            if (!hasData)
                                continue;

                            // Bỏ qua summary rows (không có SHOP = không phải data row)
                            string shopVal = "";
                            foreach (var key in new[] { "SHOP", "Shop" })
                                if (rowData.TryGetValue(key, out string sv) && !string.IsNullOrEmpty(sv))
                                { shopVal = sv; break; }
                            if (string.IsNullOrWhiteSpace(shopVal))
                                continue;

                            // Tag sheet name để user biết đơn từ ngày nào
                            rowData["_SHEET"] = ws.Name;
                            allRows.Add(rowData);
                        }
                    }
                }

                if (doiSoatHeaders == null || allRows.Count == 0)
                {
                    MessageBox.Show("Không tìm thấy dữ liệu trong file đối soát.", "Thông báo");
                    return;
                }

                // Thêm cột _SHEET vào headers nếu chưa có
                if (!doiSoatHeaders.Contains("_SHEET"))
                    doiSoatHeaders.Add("_SHEET");

                // Hiển thị dialog chọn đơn import (với auto-search nếu có MÃ)
                ShowDoiSoatSelectionDialog(doiSoatHeaders, allRows, sourceGrid, lblResult, searchMaList);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi đọc file đối soát:\n{ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Dialog hiển thị dữ liệu đối soát để user chọn rows import.
        /// Nếu searchMaList có dữ liệu → auto-check rows matching MÃ HĐ.
        /// </summary>
        private void ShowDoiSoatSelectionDialog(
            List<string> headers,
            List<Dictionary<string, string>> rows,
            DataGridView sourceGrid,
            Label lblResult,
            List<string> searchMaList = null)
        {
            using var dlg = new Form
            {
                Text = "📥 Chọn đơn từ file đối soát",
                Size = new Size(900, 500),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable,
                MinimizeBox = false,
            };

            // Search box
            var pnlSearch = new Panel { Dock = DockStyle.Top, Height = 35 };
            var lblSearch = new Label
            {
                Text = "Tìm MÃ HĐ:",
                Left = 5, Top = 8, AutoSize = true,
                Font = new Font("Arial", 9),
            };
            var txtSearch = new TextBox
            {
                Left = 80, Top = 5, Width = 200,
                Font = new Font("Arial", 10),
            };
            var btnSearch = new Button
            {
                Text = "Tìm", Left = 290, Top = 4, Width = 60, Height = 26,
            };
            pnlSearch.Controls.AddRange(new Control[] { lblSearch, txtSearch, btnSearch });

            // DataGridView
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
            };

            // Add checkbox column
            var chkCol = new DataGridViewCheckBoxColumn
            {
                Name = "chkSelect",
                HeaderText = "✓",
                Width = 30,
            };
            dgv.Columns.Add(chkCol);

            // Add data columns
            foreach (string h in headers)
                dgv.Columns.Add(h, h);

            // Populate
            foreach (var row in rows)
            {
                var vals = new List<object> { false };
                foreach (string h in headers)
                    vals.Add(row.GetValueOrDefault(h, ""));
                dgv.Rows.Add(vals.ToArray());
            }

            // Search handler
            btnSearch.Click += (s, e) =>
            {
                string search = txtSearch.Text.Trim();
                if (string.IsNullOrEmpty(search))
                    return;
                foreach (DataGridViewRow r in dgv.Rows)
                {
                    foreach (DataGridViewCell cell in r.Cells)
                    {
                        if (cell.ColumnIndex == 0) continue; // skip checkbox
                        string v = cell.Value?.ToString() ?? "";
                        if (v.Contains(search, StringComparison.OrdinalIgnoreCase))
                        {
                            r.Cells[0].Value = true; // check it
                            r.Selected = true;
                            dgv.FirstDisplayedScrollingRowIndex = r.Index;
                            break;
                        }
                    }
                }
            };

            // ── Auto-search: nếu có danh sách MÃ → tự check rows matching ──
            // Tìm cột MÃ trong grid
            int maColGridIdx = -1;
            for (int c = 1; c < dgv.Columns.Count; c++)
            {
                string h = dgv.Columns[c].HeaderText.ToUpper().Trim();
                if (h == "MÃ" || h == "MÃ HĐ" || h == "MA")
                {
                    maColGridIdx = c;
                    break;
                }
            }

            int autoChecked = 0;
            if (searchMaList != null && searchMaList.Count > 0 && maColGridIdx >= 0)
            {
                var searchSet = new HashSet<string>(searchMaList, StringComparer.OrdinalIgnoreCase);
                foreach (DataGridViewRow r in dgv.Rows)
                {
                    if (r.IsNewRow) continue;
                    string cellMa = (r.Cells[maColGridIdx].Value?.ToString() ?? "").Trim();
                    if (searchSet.Contains(cellMa))
                    {
                        r.Cells[0].Value = true;
                        r.DefaultCellStyle.BackColor = Color.FromArgb(200, 255, 200);
                        autoChecked++;
                    }
                }

                if (autoChecked > 0)
                {
                    // Scroll to first checked row
                    foreach (DataGridViewRow r in dgv.Rows)
                        if (r.Cells[0].Value is true)
                        {
                            dgv.FirstDisplayedScrollingRowIndex = r.Index;
                            break;
                        }
                }

                // Thông báo kết quả auto-search
                var notFoundMa = searchMaList.Where(m =>
                {
                    foreach (DataGridViewRow r in dgv.Rows)
                    {
                        if (r.IsNewRow) continue;
                        string cellMa = (r.Cells[maColGridIdx].Value?.ToString() ?? "").Trim();
                        if (cellMa.Equals(m, StringComparison.OrdinalIgnoreCase))
                            return false;
                    }
                    return true;
                }).ToList();

                string title = $"🔍 Tìm thấy {autoChecked}/{searchMaList.Count} đơn trong đối soát";
                if (notFoundMa.Count > 0)
                    title += $"  |  ❌ Không tìm: {string.Join(", ", notFoundMa)}";
                dlg.Text = title;
            }
            else if (searchMaList != null && searchMaList.Count > 0 && maColGridIdx < 0)
            {
                dlg.Text = "📥 Chọn đơn từ đối soát (⚠ không tìm được cột MÃ để auto-search)";
            }

            // Import button
            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 45 };
            var btnImport = new Button
            {
                Text = "📥 Import đã chọn + đánh dấu",
                Left = 10, Top = 8, Width = 250, Height = 30,
                BackColor = Color.FromArgb(40, 40, 40),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Arial", 10),
            };
            btnImport.FlatAppearance.BorderSize = 0;

            btnImport.Click += (s, e) =>
            {
                int imported = ImportSelectedDoiSoatRows(dgv, headers, sourceGrid);
                if (lblResult != null)
                    lblResult.Text = $"✅ Đã import {imported} đơn từ đối soát";
                dlg.Close();
            };

            pnlBottom.Controls.Add(btnImport);
            dlg.Controls.Add(dgv);
            dlg.Controls.Add(pnlSearch);
            dlg.Controls.Add(pnlBottom);
            dlg.ShowDialog();
        }

        /// <summary>
        /// Import các row đã check từ đối soát grid vào source grid + đánh dấu return.
        /// </summary>
        private int ImportSelectedDoiSoatRows(
            DataGridView doiSoatGrid,
            List<string> doiSoatHeaders,
            DataGridView sourceGrid)
        {
            int imported = 0;

            // Map đối soát headers → source grid column indices
            var sourceColMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = 0; c < sourceGrid.Columns.Count; c++)
            {
                string h = sourceGrid.Columns[c].HeaderText;
                sourceColMap[h] = c;
            }

            // Detect source columns for marking
            int srcColUngTien = -1, srcColHangTon = -1, srcColFail = -1;
            for (int c = 0; c < sourceGrid.Columns.Count; c++)
            {
                string h = sourceGrid.Columns[c].HeaderText.ToLower();
                if (h.Contains("ứng tiền") || h.Contains("ung tien"))
                    srcColUngTien = c;
                if (h.Contains("hàng tồn") || h.Contains("hang ton"))
                    srcColHangTon = c;
                if (h.Contains("fail"))
                    srcColFail = c;
            }

            foreach (DataGridViewRow doiSoatRow in doiSoatGrid.Rows)
            {
                if (doiSoatRow.IsNewRow)
                    continue;

                // Check if checkbox is selected
                var chkVal = doiSoatRow.Cells[0].Value;
                if (chkVal == null || !(bool)chkVal)
                    continue;

                // Create new row in source grid
                var newRow = new DataGridViewRow();
                newRow.CreateCells(sourceGrid);

                // Map đối soát data → source columns (best-effort matching by header name)
                for (int di = 0; di < doiSoatHeaders.Count; di++)
                {
                    string doiSoatHeader = doiSoatHeaders[di];
                    int dgvColIdx = di + 1; // +1 because first column is checkbox

                    // Try to find matching source column
                    int srcColIdx = FindMatchingSourceColumn(doiSoatHeader, sourceColMap);
                    if (srcColIdx < 0 || srcColIdx >= newRow.Cells.Count)
                        continue;

                    object val = doiSoatRow.Cells[dgvColIdx].Value;
                    newRow.Cells[srcColIdx].Value = val;
                }

                // Đánh dấu return (loại 2: ỨNG TIỀN=x, HÀNG TỒN=x, FAIL=xx)
                if (srcColUngTien >= 0 && srcColUngTien < newRow.Cells.Count)
                    newRow.Cells[srcColUngTien].Value = "x";
                if (srcColHangTon >= 0 && srcColHangTon < newRow.Cells.Count)
                    newRow.Cells[srcColHangTon].Value = "x";
                if (srcColFail >= 0 && srcColFail < newRow.Cells.Count)
                    newRow.Cells[srcColFail].Value = "xx";

                sourceGrid.Rows.Add(newRow);

                // Tô nền đỏ nhạt row vừa import
                int lastIdx = sourceGrid.Rows.Count - 1;
                if (!sourceGrid.Rows[lastIdx].IsNewRow)
                {
                    for (int c = 0; c < sourceGrid.Columns.Count; c++)
                        sourceGrid.Rows[lastIdx].Cells[c].Style.BackColor = Color.FromArgb(255, 220, 220);
                }

                imported++;
            }

            return imported;
        }

        /// <summary>
        /// Tìm source column index matching với đối soát header.
        /// Dùng fuzzy match: "MÃ" ↔ "MÃ", "TIỀN THU" ↔ "TIỀN THU", etc.
        /// </summary>
        private static int FindMatchingSourceColumn(
            string doiSoatHeader,
            Dictionary<string, int> sourceColMap)
        {
            // Exact match
            if (sourceColMap.TryGetValue(doiSoatHeader, out int idx))
                return idx;

            // Partial match
            string lower = doiSoatHeader.ToLower();
            foreach (var kvp in sourceColMap)
            {
                string srcLower = kvp.Key.ToLower();
                if (srcLower.Contains(lower) || lower.Contains(srcLower))
                    return kvp.Value;
            }

            // Common alias mapping
            var aliases = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                { "mã", new[] { "mã hđ", "mã hd", "số hđ", "so hd", "ma hd" } },
                { "tên kh", new[] { "tên khách", "ten kh", "khách hàng", "tên" } },
                { "tiền thu", new[] { "thu", "tổng thu", "tien thu" } },
                { "tiền ship", new[] { "ship", "phí ship", "tien ship" } },
                { "địa chỉ", new[] { "dia chi", "đ/c" } },
                { "quận", new[] { "quan", "q." } },
            };

            foreach (var kvp in aliases)
            {
                bool doiSoatMatch = lower.Contains(kvp.Key.ToLower())
                    || kvp.Value.Any(a => lower.Contains(a.ToLower()));
                if (!doiSoatMatch)
                    continue;

                // Tìm source column matching key
                if (sourceColMap.TryGetValue(kvp.Key, out int srcIdx))
                    return srcIdx;
                foreach (string alias in kvp.Value)
                    if (sourceColMap.TryGetValue(alias, out int aIdx))
                        return aIdx;
            }

            return -1;
        }

        /// <summary>
        /// Lấy DataGridView hiện tại từ tab Excel Viewer.
        /// </summary>
        private DataGridView GetCurrentSourceGrid()
        {
            if (tabExcelSheets == null || tabExcelSheets.TabPages.Count == 0)
                return null;
            var currentSheet = tabExcelSheets.SelectedTab;
            if (currentSheet == null)
                return null;
            foreach (Control ctrl in currentSheet.Controls)
                if (ctrl is DataGridView dgv)
                    return dgv;
            return null;
        }
    }
}

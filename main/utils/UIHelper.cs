using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace TextInputter
{
    /// <summary>
    /// Helper class for UI control creation and styling
    /// </summary>
    public static class UIHelper
    {
        /// <summary>
        /// Create a label-textbox pair
        /// </summary>
        public static TextBox CreateLabelTextBox(Panel panel, string labelText, ref int yPos, bool isMultiline = false)
        {
            Label lbl = new Label
            {
                Text = labelText,
                AutoSize = true,
                Location = new Point(10, yPos)
            };
            panel.Controls.Add(lbl);

            TextBox txt = new TextBox
            {
                Location = new Point(120, yPos),
                Width = 300,
                Height = isMultiline ? 60 : 22,
                Multiline = isMultiline,
                ScrollBars = isMultiline ? ScrollBars.Vertical : ScrollBars.None
            };
            panel.Controls.Add(txt);

            yPos += (isMultiline ? 75 : 35);
            return txt;
        }

        /// <summary>
        /// Create a styled button
        /// </summary>
        public static Button CreateButton(string text, Color backColor, int x, int y, int width = 100, int height = 30)
        {
            return new Button
            {
                Text = text,
                BackColor = backColor,
                ForeColor = Color.White,
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(x, y),
                Size = new Size(width, height),
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat
            };
        }

        /// <summary>
        /// Create a read-only textbox for display
        /// </summary>
        public static TextBox CreateReadOnlyTextBox(Panel panel, string labelText, ref int yPos, bool isMultiline = false)
        {
            var txt = CreateLabelTextBox(panel, labelText, ref yPos, isMultiline);
            txt.ReadOnly = true;
            txt.BackColor = Color.WhiteSmoke;
            return txt;
        }

        /// <summary>
        /// Create a styled section label
        /// </summary>
        public static Label CreateSectionLabel(Panel panel, string text, ref int yPos)
        {
            Label lbl = new Label
            {
                Text = text,
                AutoSize = true,
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(10, yPos)
            };
            panel.Controls.Add(lbl);
            yPos += 30;
            return lbl;
        }

        // ─── RichTextBox Search ────────────────────────────────────────────────

        /// <summary>
        /// Tạo search bar (🔍 TextBox + nút ▼ ▲ ✕ + label "X/Y") gắn vào panel cha.
        /// Trả về Panel để caller dùng cho responsive resize.
        /// idxHolder[0] tự quản lý vị trí match hiện tại (closure-safe, không cần ref int).
        /// </summary>
        public static Panel CreateRichTextBoxSearchBar(Panel parent, int y, Func<RichTextBox> getTarget)
        {
            int[] idxHolder = { -1 };

            var pnl = new Panel
            {
                Location    = new Point(10, y),
                Width       = parent.ClientSize.Width - 20,
                Height      = 28,
                BackColor   = Color.FromArgb(230, 240, 255),
                BorderStyle = BorderStyle.FixedSingle
            };

            var lblIcon = new Label
            {
                Text     = "🔍",
                AutoSize = true,
                Location = new Point(4, 6),
                Font     = new Font("Arial", 9)
            };

            var txtSearch = new TextBox
            {
                Location        = new Point(24, 4),
                Width           = 200,
                Height          = 20,
                BorderStyle     = BorderStyle.FixedSingle,
                Font            = new Font("Arial", 9),
                PlaceholderText = "Tìm kiếm..."
            };

            var btnNext = new Button
            {
                Text      = "▼",
                Location  = new Point(230, 3),
                Width     = 30,
                Height    = 22,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Arial", 8),
                BackColor = Color.White
            };
            btnNext.FlatAppearance.BorderSize = 1;

            var btnPrev = new Button
            {
                Text      = "▲",
                Location  = new Point(264, 3),
                Width     = 30,
                Height    = 22,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Arial", 8),
                BackColor = Color.White
            };
            btnPrev.FlatAppearance.BorderSize = 1;

            var btnClear = new Button
            {
                Text      = "✕",
                Location  = new Point(298, 3),
                Width     = 24,
                Height    = 22,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Arial", 8),
                BackColor = Color.White,
                ForeColor = Color.Red
            };
            btnClear.FlatAppearance.BorderSize = 0;

            var lblResult = new Label
            {
                Text      = "",
                AutoSize  = true,
                Location  = new Point(328, 7),
                Font      = new Font("Arial", 8),
                ForeColor = Color.DimGray
            };

            btnNext.Click += (s, e) => SearchInRichTextBox(getTarget(), txtSearch.Text, true,  idxHolder, lblResult);
            btnPrev.Click += (s, e) => SearchInRichTextBox(getTarget(), txtSearch.Text, false, idxHolder, lblResult);
            btnClear.Click += (s, e) =>
            {
                txtSearch.Clear();
                lblResult.Text = "";
                idxHolder[0]   = -1;
                ClearRichTextBoxHighlights(getTarget());
            };
            txtSearch.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; btnNext.PerformClick(); }
            };

            pnl.Controls.AddRange(new Control[] { lblIcon, txtSearch, btnNext, btnPrev, btnClear, lblResult });
            parent.Controls.Add(pnl);
            return pnl;
        }

        /// <summary>
        /// Tìm kiếm text trong RichTextBox:
        ///   - Highlight tất cả match màu vàng
        ///   - Match hiện tại highlight màu cam
        ///   - Scroll đến match hiện tại
        ///   - Cập nhật label "X/Y"
        /// </summary>
        public static void SearchInRichTextBox(RichTextBox rtb, string term, bool forward, int[] idxHolder, Label lblResult)
        {
            if (rtb == null || string.IsNullOrWhiteSpace(term))
            {
                if (lblResult != null) lblResult.Text = "";
                return;
            }

            string textLow = rtb.Text.ToLowerInvariant();
            string termLow = term.ToLowerInvariant();

            var matches = new List<int>();
            int pos = 0;
            while ((pos = textLow.IndexOf(termLow, pos)) >= 0)
            {
                matches.Add(pos);
                pos += termLow.Length;
            }

            if (matches.Count == 0)
            {
                if (lblResult != null) { lblResult.Text = "Không tìm thấy"; lblResult.ForeColor = Color.Red; }
                ClearRichTextBoxHighlights(rtb);
                idxHolder[0] = -1;
                return;
            }

            ClearRichTextBoxHighlights(rtb);
            rtb.SuspendLayout();
            foreach (int m in matches)
            {
                rtb.Select(m, term.Length);
                rtb.SelectionBackColor = Color.Yellow;
                rtb.SelectionColor     = Color.Black;
            }

            idxHolder[0] = forward
                ? (idxHolder[0] + 1) % matches.Count
                : (idxHolder[0] - 1 + matches.Count) % matches.Count;

            int cur = matches[idxHolder[0]];
            rtb.Select(cur, term.Length);
            rtb.SelectionBackColor = Color.Orange;
            rtb.SelectionColor     = Color.Black;

            rtb.ResumeLayout();
            rtb.ScrollToCaret();

            if (lblResult != null) { lblResult.Text = $"{idxHolder[0] + 1}/{matches.Count}"; lblResult.ForeColor = Color.DarkGreen; }
        }

        /// <summary>
        /// Xóa toàn bộ highlight trong RichTextBox (reset về nền trắng, chữ mặc định).
        /// </summary>
        public static void ClearRichTextBoxHighlights(RichTextBox rtb)
        {
            if (rtb == null || rtb.TextLength == 0) return;
            int savedStart  = rtb.SelectionStart;
            int savedLength = rtb.SelectionLength;
            rtb.SuspendLayout();
            rtb.SelectAll();
            rtb.SelectionBackColor = Color.White;
            rtb.SelectionColor     = rtb.ForeColor;
            rtb.Select(savedStart, savedLength);
            rtb.ResumeLayout();
        }
    }
}

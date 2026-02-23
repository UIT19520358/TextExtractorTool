using System;
using System.Drawing;
using System.Windows.Forms;

namespace TextInputter.Services
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
    }
}

using System.Drawing;
using System.Windows.Forms;

namespace TextInputter
{
    /// <summary>
    /// InvoiceTab UI declarations — control fields + InitializeInvoiceTabUI().
    /// Logic ở InvoiceTab.cs.
    /// </summary>
    public partial class MainForm
    {
        // ─── Controls thuộc Excel Viewer Tab ──────────────────────────────────
        private TabControl tabExcelSheets;
        private Panel panelExcelButtons;
        private Button btnSaveExcelEditor;
        private Button btnUndoExcelEditor;
        private Button btnCancelExcelEditor;
        private Button btnCalculateExcelData;

        // ─── Controls thuộc Invoice / Daily Report Tab ─────────────────────────
        private DataGridView dgvInvoice;
        private Label lblInvoiceTotal;

        // ─── Controls tham chiếu (ẩn, giữ để tránh lỗi wire Designer) ────────
        private Label lblInvoiceDate;


        // ─── State ────────────────────────────────────────────────────────────
        private string currentExcelFilePath;

        /// <summary>
        /// Khởi tạo toàn bộ UI cho tabExcelViewer + tabInvoice.
        /// Gọi từ MainForm constructor sau InitializeComponent().
        /// </summary>
        private void InitializeInvoiceTabUI()
        {
            // ── Instantiate controls ──────────────────────────────────────────
            tabExcelSheets = new TabControl();
            panelExcelButtons = new Panel();
            btnSaveExcelEditor = new Button();
            btnUndoExcelEditor = new Button();
            btnCancelExcelEditor = new Button();
            btnCalculateExcelData = new Button();

            dgvInvoice = new DataGridView();
            lblInvoiceTotal = new Label();
            lblInvoiceDate = new Label();


            // ── tabExcelViewer layout ─────────────────────────────────────────

            // panelExcelButtons (toolbar ở trên)
            panelExcelButtons.BackColor = System.Drawing.Color.White;
            panelExcelButtons.Dock = DockStyle.Top;
            panelExcelButtons.Height = 35;
            panelExcelButtons.Name = "panelExcelButtons";
            panelExcelButtons.Padding = new System.Windows.Forms.Padding(5);
            panelExcelButtons.TabIndex = 0;

            // btnSaveExcelEditor
            btnSaveExcelEditor.BackColor = System.Drawing.Color.FromArgb(40, 40, 40);
            btnSaveExcelEditor.FlatStyle = FlatStyle.Flat;
            btnSaveExcelEditor.FlatAppearance.BorderSize = 0;
            btnSaveExcelEditor.ForeColor = System.Drawing.Color.White;
            btnSaveExcelEditor.Location = new System.Drawing.Point(5, 5);
            btnSaveExcelEditor.Name = "btnSaveExcelEditor";
            btnSaveExcelEditor.Size = new System.Drawing.Size(70, 25);
            btnSaveExcelEditor.Text = "💾 Lưu";
            btnSaveExcelEditor.Font = new System.Drawing.Font("Arial", 9F);
            btnSaveExcelEditor.Click += BtnSaveExcelEditor_Click;

            // btnUndoExcelEditor
            btnUndoExcelEditor.BackColor = System.Drawing.Color.FromArgb(40, 40, 40);
            btnUndoExcelEditor.FlatStyle = FlatStyle.Flat;
            btnUndoExcelEditor.FlatAppearance.BorderSize = 0;
            btnUndoExcelEditor.ForeColor = System.Drawing.Color.White;
            btnUndoExcelEditor.Location = new System.Drawing.Point(80, 5);
            btnUndoExcelEditor.Name = "btnUndoExcelEditor";
            btnUndoExcelEditor.Size = new System.Drawing.Size(70, 25);
            btnUndoExcelEditor.Text = "↶ Undo";
            btnUndoExcelEditor.Font = new System.Drawing.Font("Arial", 9F);
            btnUndoExcelEditor.Click += BtnUndoExcelEditor_Click;

            // btnCancelExcelEditor
            btnCancelExcelEditor.BackColor = System.Drawing.Color.FromArgb(40, 40, 40);
            btnCancelExcelEditor.FlatStyle = FlatStyle.Flat;
            btnCancelExcelEditor.FlatAppearance.BorderSize = 0;
            btnCancelExcelEditor.ForeColor = System.Drawing.Color.White;
            btnCancelExcelEditor.Location = new System.Drawing.Point(155, 5);
            btnCancelExcelEditor.Name = "btnCancelExcelEditor";
            btnCancelExcelEditor.Size = new System.Drawing.Size(70, 25);
            btnCancelExcelEditor.Text = "✕ Đóng";
            btnCancelExcelEditor.Font = new System.Drawing.Font("Arial", 9F);
            btnCancelExcelEditor.Click += BtnCancelExcelEditor_Click;

            // btnCalculateExcelData
            btnCalculateExcelData.BackColor = System.Drawing.Color.FromArgb(40, 40, 40);
            btnCalculateExcelData.FlatStyle = FlatStyle.Flat;
            btnCalculateExcelData.FlatAppearance.BorderSize = 0;
            btnCalculateExcelData.ForeColor = System.Drawing.Color.White;
            btnCalculateExcelData.Location = new System.Drawing.Point(230, 5);
            btnCalculateExcelData.Name = "btnCalculateExcelData";
            btnCalculateExcelData.Size = new System.Drawing.Size(90, 25);
            btnCalculateExcelData.Text = "🧮 Tính Tiền";
            btnCalculateExcelData.Font = new System.Drawing.Font("Arial", 9F);
            btnCalculateExcelData.Click += BtnCalculateExcelData_Click;

            // btnMarkReturns — nút đánh dấu đơn trả
            var btnMarkReturns = new Button
            {
                BackColor = System.Drawing.Color.FromArgb(40, 40, 40),
                FlatStyle = FlatStyle.Flat,
                ForeColor = System.Drawing.Color.White,
                Location = new System.Drawing.Point(325, 5),
                Name = "btnMarkReturns",
                Size = new System.Drawing.Size(100, 25),
                Text = "↩ Đơn Trả",
                Font = new System.Drawing.Font("Arial", 9F),
            };
            btnMarkReturns.FlatAppearance.BorderSize = 0;
            btnMarkReturns.Click += (s, e) => ShowReturnDialog();

            panelExcelButtons.Controls.Add(btnSaveExcelEditor);
            panelExcelButtons.Controls.Add(btnUndoExcelEditor);
            panelExcelButtons.Controls.Add(btnCancelExcelEditor);
            panelExcelButtons.Controls.Add(btnCalculateExcelData);
            panelExcelButtons.Controls.Add(btnMarkReturns);

            // tabExcelSheets (fill phần còn lại bên dưới toolbar)
            tabExcelSheets.Dock = DockStyle.Fill;
            tabExcelSheets.Name = "tabExcelSheets";
            tabExcelSheets.SelectedIndex = 0;
            tabExcelSheets.TabIndex = 1;

            tabExcelViewer.Controls.Add(tabExcelSheets);
            tabExcelViewer.Controls.Add(panelExcelButtons);

            // ── tabInvoice layout ─────────────────────────────────────────────

            dgvInvoice.BackgroundColor = System.Drawing.Color.White;
            dgvInvoice.ColumnHeadersHeightSizeMode =
                DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvInvoice.Dock = DockStyle.Fill;
            dgvInvoice.Name = "dgvInvoice";
            dgvInvoice.TabIndex = 0;
            dgvInvoice.ScrollBars = ScrollBars.Both;
            dgvInvoice.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            lblInvoiceTotal.AutoSize = false;
            lblInvoiceTotal.Height = 40;
            lblInvoiceTotal.Font = new System.Drawing.Font(
                "Arial",
                11F,
                System.Drawing.FontStyle.Bold
            );
            lblInvoiceTotal.ForeColor = System.Drawing.Color.FromArgb(40, 40, 40);
            lblInvoiceTotal.BackColor = System.Drawing.Color.LightYellow;
            lblInvoiceTotal.Name = "lblInvoiceTotal";
            lblInvoiceTotal.Text = "TỔNG CỘNG: 0 đ";
            lblInvoiceTotal.Dock = DockStyle.Top;
            lblInvoiceTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            lblInvoiceTotal.Padding = new Padding(10, 0, 0, 0);
            lblInvoiceTotal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;

            tabInvoice.Controls.Add(lblInvoiceTotal);
            tabInvoice.Controls.Add(dgvInvoice);

        }
    }
}

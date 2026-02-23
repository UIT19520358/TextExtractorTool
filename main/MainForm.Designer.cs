using System.Drawing;
using System.Windows.Forms;

namespace TextInputter
{
    partial class MainForm
    {
    private Panel panelTop;
    private Panel panelLeft;
        private Panel panelBottom;
        private Label lblTitle;
        private Button btnSelectFolder;
        private Button btnStart;
    private Button btnOpenExcel;
        private Button btnClear;
        private Button btnExit;
        private Label lblFolderPath;
        private Label lblImageCount;
        private Label lblStatus;
        private Label lblCurrentFile;
        private RichTextBox txtResult;
        private ProgressBar progressBar;
        private TabControl tabMainControl;
        private TabPage tabOCR;
        private TabPage tabExcelViewer;
        private TabPage tabInvoice;
        private TabPage tabManualInput;
        private TabControl tabExcelSheets;
        private DataGridView dgvInvoice;
        private Label lblInvoiceDate;
        private Button btnAddInvoiceRow;
        private Button btnSaveInvoice;
        private Button btnImportFromExcel;
        private Button btnCalculateInvoice;
        private Label lblInvoiceTotal;
        private string currentExcelFilePath;
        private Button btnSaveExcelEditor;
        private Button btnUndoExcelEditor;
        private Button btnCancelExcelEditor;
        private Panel panelExcelButtons;
        private Button btnCalculateExcelData;

        private void InitializeComponent()
        {
            panelTop = new Panel();
            panelLeft = new Panel();
            panelBottom = new Panel();
            lblTitle = new Label();
            btnSelectFolder = new Button();
            btnStart = new Button();
            btnOpenExcel = new Button();
            btnClear = new Button();
            btnExit = new Button();
            lblFolderPath = new Label();
            lblImageCount = new Label();
            lblStatus = new Label();
            lblCurrentFile = new Label();
            txtResult = new RichTextBox();
            progressBar = new ProgressBar();
            tabMainControl = new TabControl();
            tabOCR = new TabPage();
            tabExcelViewer = new TabPage();
            tabInvoice = new TabPage();
            tabManualInput = new TabPage();
            tabExcelSheets = new TabControl();
            dgvInvoice = new DataGridView();
            lblInvoiceDate = new Label();
            btnAddInvoiceRow = new Button();
            btnSaveInvoice = new Button();
            btnImportFromExcel = new Button();
            btnCalculateInvoice = new Button();
            lblInvoiceTotal = new Label();
            btnSaveExcelEditor = new Button();
            btnUndoExcelEditor = new Button();
            btnCancelExcelEditor = new Button();
            panelExcelButtons = new Panel();
            btnCalculateExcelData = new Button();

            panelTop.SuspendLayout();
            panelLeft.SuspendLayout();
            panelBottom.SuspendLayout();
            SuspendLayout();

            // panelTop
            panelTop.BackColor = Color.FromArgb(20, 20, 20);
            panelTop.Controls.Add(lblTitle);
            panelTop.Dock = DockStyle.Top;
            panelTop.Height = 50;
            panelTop.Name = "panelTop";
            panelTop.Padding = new Padding(10);
            panelTop.TabIndex = 0;

            // lblTitle
            lblTitle.AutoSize = true;
            lblTitle.Font = new Font("Arial", 14F, FontStyle.Bold);
            lblTitle.ForeColor = Color.White;
            lblTitle.Name = "lblTitle";
            lblTitle.Text = "Text Extraction Tool";
            lblTitle.Location = new Point(10, 15);
            lblTitle.TabIndex = 0;

            // panelLeft
            panelLeft.BackColor = Color.White;
            panelLeft.Controls.Add(btnSelectFolder);
            panelLeft.Controls.Add(btnStart);
            panelLeft.Controls.Add(btnOpenExcel);
            panelLeft.Controls.Add(btnClear);
            panelLeft.Controls.Add(btnExit);
            panelLeft.Controls.Add(lblFolderPath);
            panelLeft.Controls.Add(lblImageCount);
            panelLeft.Dock = DockStyle.Left;
            panelLeft.Width = 250;
            panelLeft.Name = "panelLeft";
            panelLeft.Padding = new Padding(10);
            panelLeft.TabIndex = 1;

            // btnSelectFolder
            btnSelectFolder.BackColor = Color.FromArgb(40, 40, 40);
            btnSelectFolder.Cursor = Cursors.Hand;
            btnSelectFolder.FlatStyle = FlatStyle.Flat;
            btnSelectFolder.FlatAppearance.BorderSize = 0;
            btnSelectFolder.ForeColor = Color.White;
            btnSelectFolder.Font = new Font("Arial", 10F, FontStyle.Bold);
            btnSelectFolder.Location = new Point(15, 20);
            btnSelectFolder.Name = "btnSelectFolder";
            btnSelectFolder.Size = new Size(220, 35);
            btnSelectFolder.TabIndex = 0;
            btnSelectFolder.Text = "📂 CHỌN THƯ MỤC";
            btnSelectFolder.Click += BtnSelectFolder_Click;

            // btnStart
            btnStart.BackColor = Color.White;
            btnStart.Cursor = Cursors.Hand;
            btnStart.Enabled = false;
            btnStart.FlatStyle = FlatStyle.Flat;
            btnStart.FlatAppearance.BorderSize = 0;
            btnStart.ForeColor = Color.Black;
            btnStart.Font = new Font("Arial", 10F, FontStyle.Bold);
            btnStart.Location = new Point(15, 65);
            btnStart.Name = "btnStart";
            btnStart.Size = new Size(220, 35);
            btnStart.TabIndex = 1;
            btnStart.Text = "▶ BẮT ĐẦU";
            btnStart.Click += BtnStart_Click;

            // btnOpenExcel
            btnOpenExcel.BackColor = Color.FromArgb(40, 40, 40);
            btnOpenExcel.Cursor = Cursors.Hand;
            btnOpenExcel.FlatStyle = FlatStyle.Flat;
            btnOpenExcel.FlatAppearance.BorderSize = 0;
            btnOpenExcel.ForeColor = Color.White;
            btnOpenExcel.Font = new Font("Arial", 10F, FontStyle.Bold);
            btnOpenExcel.Location = new Point(15, 110);
            btnOpenExcel.Name = "btnOpenExcel";
            btnOpenExcel.Size = new Size(220, 35);
            btnOpenExcel.TabIndex = 2;
            btnOpenExcel.Text = "📊 EXCEL";
            btnOpenExcel.Click += BtnOpenExcel_Click;

            // btnClear
            btnClear.BackColor = Color.FromArgb(40, 40, 40);
            btnClear.Cursor = Cursors.Hand;
            btnClear.FlatStyle = FlatStyle.Flat;
            btnClear.FlatAppearance.BorderSize = 0;
            btnClear.ForeColor = Color.White;
            btnClear.Font = new Font("Arial", 10F, FontStyle.Bold);
            btnClear.Location = new Point(15, 155);
            btnClear.Name = "btnClear";
            btnClear.Size = new Size(220, 35);
            btnClear.TabIndex = 3;
            btnClear.Text = "🗑 XÓA";
            btnClear.Click += BtnClear_Click;

            // btnExit
            btnExit.BackColor = Color.FromArgb(40, 40, 40);
            btnExit.Cursor = Cursors.Hand;
            btnExit.FlatStyle = FlatStyle.Flat;
            btnExit.FlatAppearance.BorderSize = 0;
            btnExit.ForeColor = Color.White;
            btnExit.Font = new Font("Arial", 10F, FontStyle.Bold);
            btnExit.Location = new Point(15, 200);
            btnExit.Name = "btnExit";
            btnExit.Size = new Size(220, 35);
            btnExit.TabIndex = 4;
            btnExit.Text = "❌ THOÁT";
            btnExit.Click += BtnExit_Click;

            // lblFolderPath
            lblFolderPath.AutoSize = true;
            lblFolderPath.Font = new Font("Arial", 8F);
            lblFolderPath.ForeColor = Color.FromArgb(100, 100, 100);
            lblFolderPath.Location = new Point(15, 240);
            lblFolderPath.Name = "lblFolderPath";
            lblFolderPath.Size = new Size(100, 13);
            lblFolderPath.TabIndex = 5;
            lblFolderPath.Text = "Chưa chọn thư mục";
            lblFolderPath.AutoEllipsis = true;
            lblFolderPath.MaximumSize = new Size(220, 40);

            // lblImageCount
            lblImageCount.AutoSize = true;
            lblImageCount.Font = new Font("Arial", 8F);
            lblImageCount.ForeColor = Color.FromArgb(100, 100, 100);
            lblImageCount.Location = new Point(15, 280);
            lblImageCount.Name = "lblImageCount";
            lblImageCount.Size = new Size(100, 13);
            lblImageCount.TabIndex = 6;
            lblImageCount.Text = "Số ảnh: 0";

            // panelBottom
            panelBottom.BackColor = Color.FromArgb(30, 30, 30);
            panelBottom.Controls.Add(lblStatus);
            panelBottom.Controls.Add(lblCurrentFile);
            panelBottom.Controls.Add(progressBar);
            panelBottom.Dock = DockStyle.Bottom;
            panelBottom.Height = 90;
            panelBottom.Name = "panelBottom";
            panelBottom.Padding = new Padding(10);
            panelBottom.TabIndex = 2;

            // lblStatus
            lblStatus.AutoSize = true;
            lblStatus.Font = new Font("Arial", 9F);
            lblStatus.ForeColor = Color.White;
            lblStatus.Location = new Point(10, 10);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(82, 15);
            lblStatus.TabIndex = 0;
            lblStatus.Text = "Trạng thái: Sẵn sàng";

            // lblCurrentFile
            lblCurrentFile.AutoSize = true;
            lblCurrentFile.Font = new Font("Arial", 8F);
            lblCurrentFile.ForeColor = Color.FromArgb(180, 180, 180);
            lblCurrentFile.Location = new Point(10, 35);
            lblCurrentFile.Name = "lblCurrentFile";
            lblCurrentFile.Size = new Size(50, 13);
            lblCurrentFile.TabIndex = 1;
            lblCurrentFile.Text = "File hiện tại: -";
            lblCurrentFile.AutoEllipsis = true;
            lblCurrentFile.MaximumSize = new Size(500, 13);

            // progressBar
            progressBar.Location = new Point(10, 60);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(500, 18);
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.TabIndex = 2;
            progressBar.Value = 0;

            // txtResult
            txtResult.Dock = DockStyle.Fill;
            txtResult.Font = new Font("Courier New", 10F);
            txtResult.BackColor = Color.White;
            txtResult.ForeColor = Color.FromArgb(40, 40, 40);
            txtResult.Name = "txtResult";
            txtResult.ReadOnly = false;
            txtResult.TabIndex = 3;
            txtResult.AllowDrop = true;
            txtResult.DragEnter += TxtResult_DragEnter;
            txtResult.DragDrop += TxtResult_DragDrop;

            // tabMainControl
            tabMainControl.Dock = DockStyle.Fill;
            tabMainControl.Name = "tabMainControl";
            tabMainControl.SelectedIndex = 0;
            tabMainControl.TabIndex = 4;
            tabMainControl.Controls.Add(tabOCR);
            tabMainControl.Controls.Add(tabExcelViewer);
            tabMainControl.Controls.Add(tabInvoice);
            tabMainControl.Controls.Add(tabManualInput);

            // tabOCR
            tabOCR.Controls.Add(txtResult);
            tabOCR.Location = new Point(4, 24);
            tabOCR.Name = "tabOCR";
            tabOCR.Padding = new Padding(3);
            tabOCR.Size = new Size(942, 572);
            tabOCR.TabIndex = 0;
            tabOCR.Text = "📝 OCR Text";
            tabOCR.UseVisualStyleBackColor = true;

            // tabExcelViewer
            tabExcelViewer.Controls.Add(tabExcelSheets);
            tabExcelViewer.Controls.Add(panelExcelButtons);
            tabExcelViewer.Location = new Point(4, 24);
            tabExcelViewer.Name = "tabExcelViewer";
            tabExcelViewer.Padding = new Padding(3);
            tabExcelViewer.Size = new Size(942, 572);
            tabExcelViewer.TabIndex = 1;
            tabExcelViewer.Text = "📊 Excel Viewer";
            tabExcelViewer.UseVisualStyleBackColor = true;

            // panelExcelButtons
            panelExcelButtons.BackColor = Color.White;
            panelExcelButtons.Controls.Add(btnSaveExcelEditor);
            panelExcelButtons.Controls.Add(btnUndoExcelEditor);
            panelExcelButtons.Controls.Add(btnCancelExcelEditor);
            panelExcelButtons.Controls.Add(btnCalculateExcelData);
            panelExcelButtons.Dock = DockStyle.Top;
            panelExcelButtons.Height = 35;
            panelExcelButtons.Name = "panelExcelButtons";
            panelExcelButtons.Padding = new Padding(5);
            panelExcelButtons.TabIndex = 0;

            // tabExcelSheets
            tabExcelSheets.Dock = DockStyle.Fill;
            tabExcelSheets.Name = "tabExcelSheets";
            tabExcelSheets.SelectedIndex = 0;
            tabExcelSheets.TabIndex = 0;

            // tabInvoice
            tabInvoice.Location = new Point(4, 24);
            tabInvoice.Name = "tabInvoice";
            tabInvoice.Padding = new Padding(3);
            tabInvoice.Size = new Size(942, 572);
            tabInvoice.TabIndex = 2;
            tabInvoice.Text = "💰 Tính Tiền";
            tabInvoice.UseVisualStyleBackColor = true;
            tabInvoice.Controls.Add(dgvInvoice);
            tabInvoice.Controls.Add(lblInvoiceTotal);

            // tabManualInput
            tabManualInput.Location = new Point(4, 24);
            tabManualInput.Name = "tabManualInput";
            tabManualInput.Padding = new Padding(3);
            tabManualInput.Size = new Size(942, 572);
            tabManualInput.TabIndex = 3;
            tabManualInput.Text = "Manual Input";
            tabManualInput.UseVisualStyleBackColor = true;

            // dgvInvoice
            dgvInvoice.BackgroundColor = Color.White;
            dgvInvoice.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvInvoice.Dock = DockStyle.Fill;
            dgvInvoice.Location = new Point(3, 3);
            dgvInvoice.Name = "dgvInvoice";
            dgvInvoice.Size = new Size(936, 500);
            dgvInvoice.TabIndex = 0;
            dgvInvoice.ScrollBars = ScrollBars.Both;
            dgvInvoice.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            // lblInvoiceDate
            lblInvoiceDate.AutoSize = true;
            lblInvoiceDate.Font = new Font("Arial", 9F, FontStyle.Bold);
            lblInvoiceDate.Location = new Point(10, 10);
            lblInvoiceDate.Name = "lblInvoiceDate";
            lblInvoiceDate.Text = "Ngày: " + DateTime.Now.ToString("dd/MM/yyyy");

            // btnAddInvoiceRow - HIDDEN (sẽ tạo trong Panel nếu cần)
            btnAddInvoiceRow.BackColor = Color.FromArgb(40, 40, 40);
            btnAddInvoiceRow.FlatStyle = FlatStyle.Flat;
            btnAddInvoiceRow.FlatAppearance.BorderSize = 0;
            btnAddInvoiceRow.ForeColor = Color.White;
            btnAddInvoiceRow.Location = new Point(800, 10);
            btnAddInvoiceRow.Name = "btnAddInvoiceRow";
            btnAddInvoiceRow.Size = new Size(60, 25);
            btnAddInvoiceRow.Text = "➕ Thêm";
            btnAddInvoiceRow.Visible = false;
            btnAddInvoiceRow.Click += BtnAddInvoiceRow_Click;

            // btnSaveInvoice - HIDDEN
            btnSaveInvoice.BackColor = Color.FromArgb(40, 40, 40);
            btnSaveInvoice.FlatStyle = FlatStyle.Flat;
            btnSaveInvoice.FlatAppearance.BorderSize = 0;
            btnSaveInvoice.ForeColor = Color.White;
            btnSaveInvoice.Location = new Point(870, 10);
            btnSaveInvoice.Name = "btnSaveInvoice";
            btnSaveInvoice.Size = new Size(60, 25);
            btnSaveInvoice.Text = "💾 Lưu";
            btnSaveInvoice.Visible = false;
            btnSaveInvoice.Click += BtnSaveInvoice_Click;

            // btnImportFromExcel - HIDDEN
            btnImportFromExcel.BackColor = Color.FromArgb(40, 40, 40);
            btnImportFromExcel.FlatStyle = FlatStyle.Flat;
            btnImportFromExcel.FlatAppearance.BorderSize = 0;
            btnImportFromExcel.ForeColor = Color.White;
            btnImportFromExcel.Location = new Point(700, 10);
            btnImportFromExcel.Name = "btnImportFromExcel";
            btnImportFromExcel.Size = new Size(95, 25);
            btnImportFromExcel.Text = "📥 Import";
            btnImportFromExcel.Visible = false;
            btnImportFromExcel.Click += BtnImportFromExcel_Click;

            // btnCalculateInvoice - HIDDEN
            btnCalculateInvoice.BackColor = Color.FromArgb(40, 40, 40);
            btnCalculateInvoice.FlatStyle = FlatStyle.Flat;
            btnCalculateInvoice.FlatAppearance.BorderSize = 0;
            btnCalculateInvoice.ForeColor = Color.White;
            btnCalculateInvoice.Location = new Point(600, 10);
            btnCalculateInvoice.Name = "btnCalculateInvoice";
            btnCalculateInvoice.Size = new Size(95, 25);
            btnCalculateInvoice.Text = "🧮 Tính Tiền";
            btnCalculateInvoice.Visible = false;
            btnCalculateInvoice.Click += BtnCalculateInvoice_Click;

            // lblInvoiceTotal
            lblInvoiceTotal.AutoSize = true;
            lblInvoiceTotal.Font = new Font("Arial", 11F, FontStyle.Bold);
            lblInvoiceTotal.ForeColor = Color.FromArgb(40, 40, 40);
            lblInvoiceTotal.Location = new Point(10, 545);
            lblInvoiceTotal.Name = "lblInvoiceTotal";
            lblInvoiceTotal.Text = "TỔNG CỘNG: 0 đ";
            lblInvoiceTotal.Dock = DockStyle.Bottom;

            // btnSaveExcelEditor
            btnSaveExcelEditor.BackColor = Color.FromArgb(40, 40, 40);
            btnSaveExcelEditor.FlatStyle = FlatStyle.Flat;
            btnSaveExcelEditor.FlatAppearance.BorderSize = 0;
            btnSaveExcelEditor.ForeColor = Color.White;
            btnSaveExcelEditor.Location = new Point(5, 5);
            btnSaveExcelEditor.Name = "btnSaveExcelEditor";
            btnSaveExcelEditor.Size = new Size(70, 25);
            btnSaveExcelEditor.Text = "💾 Lưu";
            btnSaveExcelEditor.Font = new Font("Arial", 9F);
            btnSaveExcelEditor.Click += BtnSaveExcelEditor_Click;

            // btnUndoExcelEditor
            btnUndoExcelEditor.BackColor = Color.FromArgb(40, 40, 40);
            btnUndoExcelEditor.FlatStyle = FlatStyle.Flat;
            btnUndoExcelEditor.FlatAppearance.BorderSize = 0;
            btnUndoExcelEditor.ForeColor = Color.White;
            btnUndoExcelEditor.Location = new Point(80, 5);
            btnUndoExcelEditor.Name = "btnUndoExcelEditor";
            btnUndoExcelEditor.Size = new Size(70, 25);
            btnUndoExcelEditor.Text = "↶ Undo";
            btnUndoExcelEditor.Font = new Font("Arial", 9F);
            btnUndoExcelEditor.Click += BtnUndoExcelEditor_Click;

            // btnCancelExcelEditor
            btnCancelExcelEditor.BackColor = Color.FromArgb(40, 40, 40);
            btnCancelExcelEditor.FlatStyle = FlatStyle.Flat;
            btnCancelExcelEditor.FlatAppearance.BorderSize = 0;
            btnCancelExcelEditor.ForeColor = Color.White;
            btnCancelExcelEditor.Location = new Point(155, 5);
            btnCancelExcelEditor.Name = "btnCancelExcelEditor";
            btnCancelExcelEditor.Size = new Size(70, 25);
            btnCancelExcelEditor.Text = "✕ Đóng";
            btnCancelExcelEditor.Font = new Font("Arial", 9F);
            btnCancelExcelEditor.Click += BtnCancelExcelEditor_Click;

            // btnCalculateExcelData
            btnCalculateExcelData.BackColor = Color.FromArgb(40, 40, 40);
            btnCalculateExcelData.FlatStyle = FlatStyle.Flat;
            btnCalculateExcelData.FlatAppearance.BorderSize = 0;
            btnCalculateExcelData.ForeColor = Color.White;
            btnCalculateExcelData.Location = new Point(230, 5);
            btnCalculateExcelData.Name = "btnCalculateExcelData";
            btnCalculateExcelData.Size = new Size(90, 25);
            btnCalculateExcelData.Text = "🧮 Tính Tiền";
            btnCalculateExcelData.Font = new Font("Arial", 9F);
            btnCalculateExcelData.Click += BtnCalculateExcelData_Click;

            // tabExcelSheets
            tabExcelSheets.Dock = DockStyle.Fill;
            tabExcelSheets.Name = "tabExcelSheets";
            tabExcelSheets.SelectedIndex = 0;
            tabExcelSheets.TabIndex = 0;

            // MainForm
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1200, 700);
            WindowState = FormWindowState.Maximized;
            Controls.Add(tabMainControl);
            Controls.Add(panelLeft);
            Controls.Add(panelTop);
            Controls.Add(panelBottom);
            Name = "MainForm";
            Text = "Text Extraction Tool";
            
            // Load application icon
            try
            {
                string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "resources", "app.ico");
                if (File.Exists(iconPath))
                {
                    Icon = new Icon(iconPath);
                }
            }
            catch { /* Icon loading failed, use default */ }

            panelTop.ResumeLayout(false);
            panelTop.PerformLayout();
            panelLeft.ResumeLayout(false);
            panelLeft.PerformLayout();
            panelBottom.ResumeLayout(false);
            panelBottom.PerformLayout();
            ResumeLayout(false);
        }

        // Event handler declarations
        private void BtnSelectFolder_Click(object? sender, System.EventArgs e) => btnSelectFolder_Click(sender, e);
        private void BtnStart_Click(object? sender, System.EventArgs e) => btnStart_Click(sender, e);
        private void BtnClear_Click(object? sender, System.EventArgs e) => btnClear_Click(sender, e);
        private void BtnExit_Click(object? sender, System.EventArgs e) => btnExit_Click(sender, e);
        private void TxtResult_DragEnter(object? sender, System.Windows.Forms.DragEventArgs e) => txtResult_DragEnter(sender, e);
        private void TxtResult_DragDrop(object? sender, System.Windows.Forms.DragEventArgs e) => txtResult_DragDrop(sender, e);
    }
}

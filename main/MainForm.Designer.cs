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
        private Button btnSaveToFile;
        private Button btnPrint;
        private Button btnClear;
        private Button btnExit;
        private Label lblFolderPath;
        private Label lblImageCount;
        private Label lblStatus;
        private Label lblCurrentFile;
        private RichTextBox txtResult;
        private ProgressBar progressBar;

        private void InitializeComponent()
        {
            panelTop = new Panel();
            panelLeft = new Panel();
            panelBottom = new Panel();
            lblTitle = new Label();
            btnSelectFolder = new Button();
            btnStart = new Button();
            btnSaveToFile = new Button();
            btnPrint = new Button();
            btnClear = new Button();
            btnExit = new Button();
            lblFolderPath = new Label();
            lblImageCount = new Label();
            lblStatus = new Label();
            lblCurrentFile = new Label();
            txtResult = new RichTextBox();
            progressBar = new ProgressBar();

            panelTop.SuspendLayout();
            panelLeft.SuspendLayout();
            panelBottom.SuspendLayout();
            SuspendLayout();

            // panelTop
            panelTop.BackColor = Color.FromArgb(41, 128, 185);
            panelTop.Controls.Add(lblTitle);
            panelTop.Dock = DockStyle.Top;
            panelTop.Height = 50;
            panelTop.Name = "panelTop";
            panelTop.Padding = new Padding(10);
            panelTop.TabIndex = 0;

            // lblTitle
            lblTitle.AutoSize = true;
            lblTitle.Font = new Font("Arial", 12F, FontStyle.Bold);
            lblTitle.ForeColor = Color.White;
            lblTitle.Name = "lblTitle";
            lblTitle.Text = "Vietnamese Text Extraction Tool";
            lblTitle.Location = new Point(10, 15);
            lblTitle.TabIndex = 0;

            // panelLeft
            panelLeft.BackColor = Color.FromArgb(236, 236, 236);
            panelLeft.Controls.Add(btnSelectFolder);
            panelLeft.Controls.Add(btnStart);
            panelLeft.Controls.Add(btnSaveToFile);
            panelLeft.Controls.Add(btnPrint);
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
            btnSelectFolder.BackColor = Color.FromArgb(41, 128, 185);
            btnSelectFolder.Cursor = Cursors.Hand;
            btnSelectFolder.FlatStyle = FlatStyle.Flat;
            btnSelectFolder.ForeColor = Color.White;
            btnSelectFolder.Font = new Font("Arial", 9F, FontStyle.Bold);
            btnSelectFolder.Location = new Point(15, 20);
            btnSelectFolder.Name = "btnSelectFolder";
            btnSelectFolder.Size = new Size(220, 30);
            btnSelectFolder.TabIndex = 0;
            btnSelectFolder.Text = "📂 CHỌN THƯ MỤC";
            btnSelectFolder.Click += BtnSelectFolder_Click;

            // btnStart
            btnStart.BackColor = Color.FromArgb(39, 174, 96);
            btnStart.Cursor = Cursors.Hand;
            btnStart.Enabled = false;
            btnStart.FlatStyle = FlatStyle.Flat;
            btnStart.ForeColor = Color.White;
            btnStart.Font = new Font("Arial", 9F, FontStyle.Bold);
            btnStart.Location = new Point(15, 60);
            btnStart.Name = "btnStart";
            btnStart.Size = new Size(220, 30);
            btnStart.TabIndex = 1;
            btnStart.Text = "▶ BẮT ĐẦU";
            btnStart.Click += BtnStart_Click;

            // btnSaveToFile
            btnSaveToFile.BackColor = Color.FromArgb(155, 89, 182);
            btnSaveToFile.Cursor = Cursors.Hand;
            btnSaveToFile.FlatStyle = FlatStyle.Flat;
            btnSaveToFile.ForeColor = Color.White;
            btnSaveToFile.Font = new Font("Arial", 9F, FontStyle.Bold);
            btnSaveToFile.Location = new Point(15, 100);
            btnSaveToFile.Name = "btnSaveToFile";
            btnSaveToFile.Size = new Size(220, 30);
            btnSaveToFile.TabIndex = 2;
            btnSaveToFile.Text = "💾 LƯU VÀO FILE";
            btnSaveToFile.Click += BtnSaveToFile_Click;

            // btnPrint
            btnPrint.BackColor = Color.FromArgb(230, 126, 34);
            btnPrint.Cursor = Cursors.Hand;
            btnPrint.FlatStyle = FlatStyle.Flat;
            btnPrint.ForeColor = Color.White;
            btnPrint.Font = new Font("Arial", 9F, FontStyle.Bold);
            btnPrint.Location = new Point(15, 140);
            btnPrint.Name = "btnPrint";
            btnPrint.Size = new Size(220, 30);
            btnPrint.TabIndex = 3;
            btnPrint.Text = "🖨 IN";
            btnPrint.Click += BtnPrint_Click;

            // btnClear
            btnClear.BackColor = Color.FromArgb(231, 76, 60);
            btnClear.Cursor = Cursors.Hand;
            btnClear.FlatStyle = FlatStyle.Flat;
            btnClear.ForeColor = Color.White;
            btnClear.Font = new Font("Arial", 9F, FontStyle.Bold);
            btnClear.Location = new Point(15, 180);
            btnClear.Name = "btnClear";
            btnClear.Size = new Size(220, 30);
            btnClear.TabIndex = 4;
            btnClear.Text = "🗑 XÓA";
            btnClear.Click += BtnClear_Click;

            // btnExit
            btnExit.BackColor = Color.FromArgb(149, 165, 166);
            btnExit.Cursor = Cursors.Hand;
            btnExit.FlatStyle = FlatStyle.Flat;
            btnExit.ForeColor = Color.White;
            btnExit.Font = new Font("Arial", 9F, FontStyle.Bold);
            btnExit.Location = new Point(15, 220);
            btnExit.Name = "btnExit";
            btnExit.Size = new Size(220, 30);
            btnExit.TabIndex = 5;
            btnExit.Text = "❌ THOÁT";
            btnExit.Click += BtnExit_Click;

            // lblFolderPath
            lblFolderPath.AutoSize = true;
            lblFolderPath.Font = new Font("Arial", 8F);
            lblFolderPath.ForeColor = Color.FromArgb(127, 140, 141);
            lblFolderPath.Location = new Point(15, 270);
            lblFolderPath.Name = "lblFolderPath";
            lblFolderPath.Size = new Size(100, 13);
            lblFolderPath.TabIndex = 6;
            lblFolderPath.Text = "Chưa chọn thư mục";
            lblFolderPath.AutoEllipsis = true;
            lblFolderPath.MaximumSize = new Size(220, 40);

            // lblImageCount
            lblImageCount.AutoSize = true;
            lblImageCount.Font = new Font("Arial", 8F);
            lblImageCount.ForeColor = Color.FromArgb(127, 140, 141);
            lblImageCount.Location = new Point(15, 320);
            lblImageCount.Name = "lblImageCount";
            lblImageCount.Size = new Size(100, 13);
            lblImageCount.TabIndex = 7;
            lblImageCount.Text = "Số ảnh: 0";

            // panelBottom
            panelBottom.BackColor = Color.FromArgb(44, 62, 80);
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
            lblCurrentFile.ForeColor = Color.FromArgb(189, 195, 199);
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
            txtResult.BackColor = Color.FromArgb(240, 240, 240);
            txtResult.Name = "txtResult";
            txtResult.ReadOnly = false;
            txtResult.TabIndex = 3;
            txtResult.AllowDrop = true;
            txtResult.DragEnter += TxtResult_DragEnter;
            txtResult.DragDrop += TxtResult_DragDrop;

            // MainForm
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1200, 700);
            Controls.Add(txtResult);
            Controls.Add(panelLeft);
            Controls.Add(panelTop);
            Controls.Add(panelBottom);
            Name = "MainForm";
            Text = "Vietnamese Text Extraction Tool";
            Icon = null;
            WindowState = FormWindowState.Normal;

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
        private void BtnSaveToFile_Click(object? sender, System.EventArgs e) => btnSaveToFile_Click(sender, e);
        private void BtnPrint_Click(object? sender, System.EventArgs e) => btnPrint_Click(sender, e);
        private void BtnClear_Click(object? sender, System.EventArgs e) => btnClear_Click(sender, e);
        private void BtnExit_Click(object? sender, System.EventArgs e) => btnExit_Click(sender, e);
        private void TxtResult_DragEnter(object? sender, System.Windows.Forms.DragEventArgs e) => txtResult_DragEnter(sender, e);
        private void TxtResult_DragDrop(object? sender, System.Windows.Forms.DragEventArgs e) => txtResult_DragDrop(sender, e);
    }
}

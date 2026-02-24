using System;
using System.Windows.Forms;
using System.Drawing;

namespace TextInputter
{
    public partial class AddressParsingDialog : Form
    {
        public string SoNha { get; set; }
        public string TenDuong { get; set; }
        public string Phuong { get; set; }
        public string Quan { get; set; }

        private Label lblTitle;
        private Label lblOriginal;
        private TextBox txtOriginal;
        private Label lblSoNha;
        private TextBox txtSoNha;
        private Label lblTenDuong;
        private TextBox txtTenDuong;
        private Label lblPhuong;
        private TextBox txtPhuong;
        private Label lblQuan;
        private TextBox txtQuan;
        private Button btnOK;
        private Button btnCancel;

        public AddressParsingDialog(string originalAddress, string soNha, string tenDuong, string phuong, string quan)
        {
            InitializeComponent();
            this.Text = "Sửa Địa Chỉ";
            this.Width = 500;
            this.Height = 350;
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            CreateControls();

            // Pre-fill
            txtOriginal.Text = originalAddress;
            txtSoNha.Text = soNha;
            txtTenDuong.Text = tenDuong;
            txtPhuong.Text = phuong;
            txtQuan.Text = quan;
        }

        private void CreateControls()
        {
            int y = 20;
            const int labelWidth = 80;
            const int textBoxWidth = 380;

            // Title
            lblTitle = new Label
            {
                Text = "Vui lòng sửa địa chỉ nếu cần thiết:",
                Location = new Point(20, y),
                AutoSize = true,
                Font = new Font("Arial", 10, FontStyle.Bold)
            };
            Controls.Add(lblTitle);
            y += 30;

            // Original Address
            lblOriginal = new Label
            {
                Text = "Gốc:",
                Location = new Point(20, y),
                Width = labelWidth,
                Height = 20
            };
            Controls.Add(lblOriginal);

            txtOriginal = new TextBox
            {
                Location = new Point(110, y),
                Width = textBoxWidth,
                Height = 20,
                ReadOnly = true,
                BackColor = SystemColors.Control
            };
            Controls.Add(txtOriginal);
            y += 35;

            // Số nhà
            lblSoNha = new Label
            {
                Text = "Số nhà:",
                Location = new Point(20, y),
                Width = labelWidth,
                Height = 20
            };
            Controls.Add(lblSoNha);

            txtSoNha = new TextBox
            {
                Location = new Point(110, y),
                Width = textBoxWidth,
                Height = 20
            };
            Controls.Add(txtSoNha);
            y += 35;

            // Tên đường
            lblTenDuong = new Label
            {
                Text = "Tên đường:",
                Location = new Point(20, y),
                Width = labelWidth,
                Height = 20
            };
            Controls.Add(lblTenDuong);

            txtTenDuong = new TextBox
            {
                Location = new Point(110, y),
                Width = textBoxWidth,
                Height = 20
            };
            Controls.Add(txtTenDuong);
            y += 35;

            // Phường
            lblPhuong = new Label
            {
                Text = "Phường:",
                Location = new Point(20, y),
                Width = labelWidth,
                Height = 20
            };
            Controls.Add(lblPhuong);

            txtPhuong = new TextBox
            {
                Location = new Point(110, y),
                Width = textBoxWidth,
                Height = 20
            };
            Controls.Add(txtPhuong);
            y += 35;

            // Quận
            lblQuan = new Label
            {
                Text = "Quận:",
                Location = new Point(20, y),
                Width = labelWidth,
                Height = 20
            };
            Controls.Add(lblQuan);

            txtQuan = new TextBox
            {
                Location = new Point(110, y),
                Width = textBoxWidth,
                Height = 20
            };
            Controls.Add(txtQuan);
            y += 40;

            // Buttons
            btnOK = new Button
            {
                Text = "Lưu",
                Location = new Point(300, y),
                Width = 80,
                Height = 30,
                DialogResult = DialogResult.OK
            };
            btnOK.Click += (s, e) =>
            {
                SoNha = txtSoNha.Text.Trim();
                TenDuong = txtTenDuong.Text.Trim();
                Phuong = txtPhuong.Text.Trim();
                Quan = txtQuan.Text.Trim();
                this.DialogResult = DialogResult.OK;
                this.Close();
            };
            Controls.Add(btnOK);

            btnCancel = new Button
            {
                Text = "Hủy",
                Location = new Point(390, y),
                Width = 80,
                Height = 30,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.Click += (s, e) =>
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            };
            Controls.Add(btnCancel);
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.ResumeLayout(false);
        }
    }
}

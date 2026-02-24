using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Cloud.Vision.V1;
using TextInputter.Services;

namespace TextInputter
{
    /// <summary>
    /// MainForm — core: shared fields, constructor, shared helpers.
    /// Tab logic ở các file riêng:
    ///   tabs/OcrTab.cs          — OCR Tab
    ///   tabs/InvoiceTab.cs      — Invoice / Excel Viewer / Daily Report
    ///   tabs/ManualInputTab.cs  — Manual Input Tab
    /// </summary>
    public partial class MainForm : Form
    {
        // ─── Shared fields ─────────────────────────────────────────────────────
        private string              folderPath       = "";
        private List<string>        imageFiles       = new List<string>();
        private bool                isProcessing     = false;
        private ImageAnnotatorClient visionClient;
        private Stack<Dictionary<string, List<string[]>>> undoStack;
        private ExcelInvoiceService     _excelInvoiceService;
        private OCRTextParsingService   _ocrParsingService;
        private List<Dictionary<string, string>> mappedDataList = new List<Dictionary<string, string>>();

        // ─── Constructor ───────────────────────────────────────────────────────
        public MainForm()
        {
            InitializeComponent();

            undoStack      = new Stack<Dictionary<string, List<string[]>>>();
            mappedDataList = new List<Dictionary<string, string>>();

            InitializeServices();
            LoadApplicationIcon();

            // Init each tab (partial methods in tabs/ files)
            InitializeOCRTab();
            InitializeManualInputTab();
        }

        // ─── Service initialization ────────────────────────────────────────────
        private void InitializeServices()
        {
            try
            {
                string credPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "textinputter-4a7bda4ef67a.json");
                if (File.Exists(credPath))
                {
                    var credential = Google.Apis.Auth.OAuth2.GoogleCredential
                        .FromFile(credPath)
                        .CreateScoped(ImageAnnotatorClient.DefaultScopes);
                    visionClient = new ImageAnnotatorClientBuilder { Credential = credential }.Build();
                    Debug.WriteLine("✅ Google Vision client initialized");
                }
                else
                {
                    Debug.WriteLine($"⚠️ Google credentials file not found at: {credPath}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"⚠️ Vision client init error: {ex.Message}");
            }

            try
            {
                _excelInvoiceService = new ExcelInvoiceService();
                Debug.WriteLine("✅ ExcelInvoiceService initialized");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"⚠️ ExcelInvoiceService init error: {ex.Message}");
            }

            try
            {
                _ocrParsingService = new OCRTextParsingService();
                Debug.WriteLine("✅ OCRTextParsingService initialized");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"⚠️ OCRTextParsingService init error: {ex.Message}");
            }
        }

        // ─── Icon ─────────────────────────────────────────────────────────────
        private void LoadApplicationIcon()
        {
            try
            {
                string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "resources", "app.ico");
                if (File.Exists(iconPath))
                    this.Icon = new Icon(iconPath);
            }
            catch { /* use default */ }
        }

        // ─── Top-bar button handlers ───────────────────────────────────────────
        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            SelectOCRFolder();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (isProcessing) return;
            if (imageFiles.Count == 0)
            {
                MessageBox.Show("⚠️ Vui lòng chọn folder ảnh trước!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            isProcessing            = true;
            btnStart.Enabled        = false;
            btnSelectFolder.Enabled = false;
            btnClear.Enabled        = false;
            progressBar.Value       = 0;
            progressBar.Maximum     = imageFiles.Count;
            lblStatus.Text          = "⏳ Đang xử lý...";
            lblStatus.ForeColor     = Color.Orange;

            Task.Run(() =>
            {
                try
                {
                    ProcessImages();
                }
                catch (Exception ex)
                {
                    this.Invoke((MethodInvoker)delegate
                    {
                        MessageBox.Show($"❌ Lỗi xử lý:\n{ex.Message}", "Lỗi",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        btnStart.Enabled        = true;
                        btnSelectFolder.Enabled = true;
                        btnClear.Enabled        = true;
                        isProcessing            = false;
                        lblStatus.Text          = "❌ Lỗi";
                        lblStatus.ForeColor     = Color.Red;
                    });
                }
            });
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtResult.Clear();
            if (txtRawOCRLog  != null) txtRawOCRLog.Clear();
            if (txtProcessLog != null) txtProcessLog.Clear();
            mappedDataList.Clear();
            lblCurrentFile.Text = "";
            lblStatus.Text      = "Ready";
            lblStatus.ForeColor = SystemColors.ControlText;
            progressBar.Value   = 0;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc muốn thoát?", "Thoát",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                Application.Exit();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if (isProcessing)
            {
                e.Cancel = true;
                MessageBox.Show("⚠️ Đang xử lý, vui lòng đợi!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ─── Drag-drop (txtResult) ─────────────────────────────────────────────
        private void txtResult_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void txtResult_DragDrop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length > 0)
            {
                folderPath  = Path.GetDirectoryName(files[0]) ?? "";
                imageFiles  = GetImageFiles(folderPath);
                lblFolderPath.Text  = folderPath;
                lblImageCount.Text  = $"{imageFiles.Count} ảnh";
            }
        }

        // ─── Shared helpers ────────────────────────────────────────────────────
        private List<string> GetImageFiles(string folder)
        {
            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
                return new List<string>();
            return Directory.GetFiles(folder, "*.*")
                .Where(f => new[] { ".jpg", ".jpeg", ".png", ".bmp", ".tiff" }
                    .Contains(Path.GetExtension(f).ToLower()))
                .ToList();
        }

        /// <summary>
        /// Gọi Google Vision OCR — trả về (text, confidence).
        /// </summary>
        private (string text, float confidence) CallPythonOCR(string imagePath)
        {
            try
            {
                if (visionClient == null) return ("", 0f);

                var image    = Google.Cloud.Vision.V1.Image.FromFile(imagePath);
                var response = visionClient.Annotate(new AnnotateImageRequest
                {
                    Image    = image,
                    Features = { new Feature { Type = Feature.Types.Type.DocumentTextDetection } }
                });

                if (response?.FullTextAnnotation == null) return ("", 0f);

                float confidence = 0f;
                if (response.FullTextAnnotation.Pages?.Count > 0)
                    confidence = response.FullTextAnnotation.Pages[0].Confidence * 100f;

                string rawText = response.FullTextAnnotation.Text ?? "";
                string cleaned = CleanOCRText(rawText);
                return (cleaned, confidence);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OCR error for {imagePath}: {ex.Message}");
                return ("", 0f);
            }
        }

        private string CleanOCRText(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "";
            var lines = raw.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var sb    = new StringBuilder();
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (!string.IsNullOrWhiteSpace(trimmed) && !IsGarbageLine(trimmed))
                    sb.AppendLine(trimmed);
            }
            return sb.ToString().Trim();
        }

        private bool IsGarbageLine(string line)
        {
            if (line.Length < 2) return true;
            int alphaNum = line.Count(c => char.IsLetterOrDigit(c));
            return (double)alphaNum / line.Length < 0.3;
        }

        private System.Drawing.Bitmap PreprocessImage(string imagePath)
        {
            var bmp = new System.Drawing.Bitmap(imagePath);
            return bmp;
        }
    }
}

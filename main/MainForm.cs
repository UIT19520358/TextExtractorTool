using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using Google.Cloud.Vision.V1;

namespace TextInputter
{
    public partial class MainForm : Form
    {
        private string folderPath = "";
        private List<string> imageFiles = new List<string>();
        private bool isProcessing = false;
        private ImageAnnotatorClient visionClient;

        public MainForm()
        {
            InitializeComponent();
            InitializeTesseract();
        }

        private void InitializeTesseract()
        {
            try
            {
                // Set Google Cloud credentials từ file JSON
                string credPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "textinputter-4a7bda4ef67a.json");
                if (File.Exists(credPath))
                {
                    Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", credPath);
                    visionClient = ImageAnnotatorClient.Create();
                    lblStatus.Text = "✅ Sẵn sàng (Google Vision API)";
                    lblStatus.ForeColor = Color.Green;
                }
                else
                {
                    MessageBox.Show("Google credentials JSON not found!", "Warning");
                    lblStatus.Text = "❌ Google credentials not found";
                    lblStatus.ForeColor = Color.Red;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"InitializeTesseract error: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                MessageBox.Show($"Initialization error:\n{ex.Message}\n\n{ex.StackTrace}", "Error");
            }
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Chọn folder chứa ảnh";
                dialog.RootFolder = Environment.SpecialFolder.MyComputer;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    folderPath = dialog.SelectedPath;
                    lblFolderPath.Text = $"📁 {folderPath}";

                    imageFiles = GetImageFiles(folderPath);
                    lblImageCount.Text = $"📷 Tìm thấy {imageFiles.Count} ảnh";

                    if (imageFiles.Count > 0)
                    {
                        btnStart.Enabled = true;
                        lblStatus.Text = "✅ Sẵn sàng xử lý";
                        lblStatus.ForeColor = Color.Green;
                    }
                    else
                    {
                        lblStatus.Text = "❌ Không tìm thấy ảnh";
                        lblStatus.ForeColor = Color.Red;
                        btnStart.Enabled = false;
                    }
                }
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (imageFiles.Count == 0)
            {
                MessageBox.Show("❌ Vui lòng chọn folder trước", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            isProcessing = true;
            btnStart.Enabled = false;
            btnSelectFolder.Enabled = false;
            btnClear.Enabled = false;
            btnPrint.Enabled = false;
            btnSaveToFile.Enabled = false;

            txtResult.Clear();
            lblStatus.Text = "⏳ Đang xử lý...";
            lblStatus.ForeColor = Color.Orange;

            progressBar.Maximum = imageFiles.Count;
            progressBar.Value = 0;

            var task = Task.Run(() => ProcessImages());
        }

        // Hàm xử lý ảnh trước OCR để cải thiện chất lượng
        private Bitmap PreprocessImage(string imagePath)
        {
            try
            {
                using (Bitmap original = new Bitmap(imagePath))
                {
                    Bitmap processed = new Bitmap(original.Width, original.Height);

                    // Lấy thông tin pixel
                    for (int y = 0; y < original.Height; y++)
                    {
                        for (int x = 0; x < original.Width; x++)
                        {
                            Color pixel = original.GetPixel(x, y);

                            // Chuyển sang grayscale
                            int gray = (int)(pixel.R * 0.299 + pixel.G * 0.587 + pixel.B * 0.114);

                            // Tăng contrast (normalize)
                            int contrast = (int)((gray - 128) * 1.5 + 128);
                            contrast = Math.Max(0, Math.Min(255, contrast));

                            // Tăng độ sáng
                            int brightness = Math.Min(255, contrast + 20);

                            Color newColor = Color.FromArgb(brightness, brightness, brightness);
                            processed.SetPixel(x, y, newColor);
                        }
                    }

                    return processed;
                }
            }
            catch
            {
                return new Bitmap(imagePath);
            }
        }

        // Gọi Google Vision API OCR
        private (string text, float confidence) CallPythonOCR(string imagePath)
        {
            try
            {
                if (visionClient == null)
                {
                    System.Diagnostics.Debug.WriteLine("ERROR: visionClient is null");
                    return ("", 0);
                }

                // Load ảnh từ file
                System.Diagnostics.Debug.WriteLine($"Loading image from: {imagePath}");
                var image = Google.Cloud.Vision.V1.Image.FromFile(imagePath);
                
                System.Diagnostics.Debug.WriteLine("Calling Google Vision API...");
                var response = visionClient.DetectTextAsync(image);
                response.Wait();

                System.Diagnostics.Debug.WriteLine($"Response received, count: {response.Result?.Count}");

                if (response.Result == null || response.Result.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("No text detected");
                    return ("", 0);
                }

                // Extract text từ response
                var textAnnotation = response.Result[0];
                if (textAnnotation == null)
                {
                    System.Diagnostics.Debug.WriteLine("textAnnotation is null");
                    return ("", 0);
                }

                string text = textAnnotation.Description?.Trim() ?? "";
                System.Diagnostics.Debug.WriteLine($"Extracted text length: {text.Length}");

                if (string.IsNullOrEmpty(text))
                {
                    System.Diagnostics.Debug.WriteLine("Text is empty after extraction");
                    return ("", 0);
                }

                // Post-processing: lọc text rác
                text = CleanOCRText(text);
                System.Diagnostics.Debug.WriteLine($"After cleaning: {text.Length}");

                if (string.IsNullOrEmpty(text))
                {
                    return ("", 0);
                }

                // Google Vision không return confidence trực tiếp, set mặc định 95%
                float confidence = 95.0f;

                return (text, confidence);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Google Vision error: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                MessageBox.Show($"❌ Error: {ex.Message}", "Google Vision Error");
            }

            return ("", 0);
        }

        // Lọc text rác từ OCR
        private string CleanOCRText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var cleanLines = new List<string>();

            foreach (var line in lines)
            {
                string trimmed = line.Trim();
                
                // Skip dòng trống
                if (string.IsNullOrWhiteSpace(trimmed))
                    continue;

                // Skip dòng chỉ chứa ký tự lạ (số random, dấu gạch, v.v)
                if (IsGarbageLine(trimmed))
                    continue;

                // Skip dòng quá ngắn (< 3 ký tự) - thường là noise
                if (trimmed.Length < 3)
                    continue;

                cleanLines.Add(trimmed);
            }

            return string.Join("\n", cleanLines);
        }

        // Kiểm tra dòng có phải rác không
        private bool IsGarbageLine(string line)
        {
            // Nếu dòng chỉ chứa số, dấu gạch, ký tự lạ => rác
            int validCharCount = 0;
            int totalCharCount = 0;

            foreach (char c in line)
            {
                totalCharCount++;

                // Chữ Việt (khoảng U+0100 - U+01FF, U+1E00 - U+1EFF)
                bool isVietnamese = (c >= '\u0100' && c <= '\u01FF') || 
                                   (c >= '\u1E00' && c <= '\u1EFF');
                
                // Chữ Anh, số, dấu cách, dấu câu thông thường
                bool isEnglish = char.IsLetterOrDigit(c) || 
                                char.IsWhiteSpace(c) || 
                                c == ',' || c == '.' || c == '-' || 
                                c == '/' || c == ':' || c == ';' ||
                                c == '(' || c == ')';

                if (isVietnamese || isEnglish)
                    validCharCount++;
            }

            // Nếu < 70% ký tự hợp lệ => rác
            return validCharCount < (totalCharCount * 0.7);
        }

        private void ProcessImages()
        {
            StringBuilder allText = new StringBuilder();
            int successCount = 0;
            int failCount = 0;

            allText.AppendLine("╔════════════════════════════════════════════════════════╗");
            allText.AppendLine("║         KẾT QUẢ NHẬN DIỆN CHỮ (OCR) TIẾNG VIỆT       ║");
            allText.AppendLine("╚════════════════════════════════════════════════════════╝\n");
            allText.AppendLine($"📅 Ngày: {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
            allText.AppendLine($"📁 Folder: {folderPath}");
            allText.AppendLine($"📷 Tổng ảnh: {imageFiles.Count}");
            allText.AppendLine("\n" + new string('═', 60) + "\n");

            for (int i = 0; i < imageFiles.Count; i++)
            {
                string imagePath = imageFiles[i];
                string fileName = Path.GetFileName(imagePath);

                this.Invoke((MethodInvoker)delegate
                {
                    progressBar.Value = i + 1;
                    lblCurrentFile.Text = $"🔄 [{i + 1}/{imageFiles.Count}] {fileName}";
                });

                try
                {
                    // Dùng PaddleOCR (tốt hơn cho tiếng Việt)
                    var (text, confidence) = CallPythonOCR(imagePath);

                    allText.AppendLine($"\n✅ TỆP #{i + 1}: {fileName}");
                    allText.AppendLine($"   📊 Độ tin cậy: {confidence:F1}%");
                    allText.AppendLine($"   ⏱️  Thời gian: {DateTime.Now:HH:mm:ss}");
                    allText.AppendLine(new string('─', 60));

                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        allText.AppendLine("\n" + text.Trim());
                        successCount++;
                    }
                    else
                    {
                        allText.AppendLine("   ⚠️  Không nhận diện được text từ ảnh này");
                        failCount++;
                    }

                    allText.AppendLine("\n" + new string('═', 60));
                }
                catch (Exception ex)
                {
                    allText.AppendLine($"\n❌ TỆP #{i + 1}: {fileName}");
                    allText.AppendLine($"   🔴 Lỗi: {ex.Message}");
                    allText.AppendLine(new string('─', 60));
                    failCount++;
                }
            }

            allText.AppendLine("\n\n╔════════════════════════════════════════════════════════╗");
            allText.AppendLine("║                    TÓM TẮT KẾT QUẢ                      ║");
            allText.AppendLine("╚════════════════════════════════════════════════════════╝\n");
            allText.AppendLine($"✅ Thành công: {successCount}/{imageFiles.Count} ảnh");
            allText.AppendLine($"❌ Thất bại: {failCount}/{imageFiles.Count} ảnh");
            allText.AppendLine($"⏱️  Thời gian xử lý: {DateTime.Now:HH:mm:ss}\n");

            this.Invoke((MethodInvoker)delegate
            {
                txtResult.Text = allText.ToString();
                lblCurrentFile.Text = $"✅ Hoàn thành: {successCount} thành công, {failCount} thất bại";
                lblStatus.Text = "✅ Xử lý xong";
                lblStatus.ForeColor = Color.Green;

                btnStart.Enabled = true;
                btnSelectFolder.Enabled = true;
                btnClear.Enabled = true;
                btnPrint.Enabled = !string.IsNullOrEmpty(txtResult.Text);
                btnSaveToFile.Enabled = !string.IsNullOrEmpty(txtResult.Text);

                isProcessing = false;

                txtResult.SelectionStart = 0;
                txtResult.ScrollToCaret();
            });
        }

        private List<string> GetImageFiles(string folderPath)
        {
            var extensions = new[] { ".jpg", ".jpeg", ".png", ".bmp" };
            var files = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories)
                .Where(f => extensions.Contains(Path.GetExtension(f).ToLower()))
                .OrderBy(f => f)
                .ToList();

            return files;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtResult.Text))
            {
                MessageBox.Show("❌ Chưa có dữ liệu để in", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string tempFile = Path.Combine(Path.GetTempPath(), "ocr_output.txt");
                File.WriteAllText(tempFile, txtResult.Text, Encoding.UTF8);

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "notepad.exe",
                    Arguments = tempFile
                });

                MessageBox.Show("✅ Mở Notepad thành công!\n\nNhấn Ctrl+P để in.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSaveToFile_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtResult.Text))
            {
                MessageBox.Show("❌ Chưa có dữ liệu để lưu", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                dialog.DefaultExt = "txt";
                dialog.FileName = $"ocr_result_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.txt";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        File.WriteAllText(dialog.FileName, txtResult.Text, Encoding.UTF8);
                        MessageBox.Show($"✅ Lưu file thành công!\n\n{dialog.FileName}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"❌ Lỗi lưu file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtResult.Clear();
            lblFolderPath.Text = "📁 Chưa chọn folder";
            lblImageCount.Text = "📷 0 ảnh";
            lblCurrentFile.Text = "";
            progressBar.Value = 0;
            lblStatus.Text = "⏳ Chờ lệnh";
            lblStatus.ForeColor = Color.Gray;
            btnStart.Enabled = false;
            btnPrint.Enabled = false;
            btnSaveToFile.Enabled = false;
            folderPath = "";
            imageFiles.Clear();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            if (isProcessing)
            {
                MessageBox.Show("⏳ Đang xử lý, vui lòng chờ", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            this.Close();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
        }

        private void txtResult_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data!.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void txtResult_DragDrop(object sender, DragEventArgs e)
        {
            string[]? files = e.Data?.GetData(DataFormats.FileDrop) as string[];
            if (files != null && files.Length > 0)
            {
                string path = files[0];
                if (Directory.Exists(path))
                {
                    folderPath = path;
                    lblFolderPath.Text = $"📁 {folderPath}";
                    imageFiles = GetImageFiles(folderPath);
                    lblImageCount.Text = $"📷 Tìm thấy {imageFiles.Count} ảnh";

                    if (imageFiles.Count > 0)
                    {
                        btnStart.Enabled = true;
                        lblStatus.Text = "✅ Sẵn sàng xử lý";
                        lblStatus.ForeColor = Color.Green;
                    }
                }
            }
        }

        private string FixSpelling(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            return text;
        }
    }
}

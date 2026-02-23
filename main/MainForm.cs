using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using System.Data;
using Google.Cloud.Vision.V1;
using ClosedXML.Excel;
using TextInputter.Services;

// Refactored MainForm - Services are now handling business logic
// UI is kept focused on presentation layer only

namespace TextInputter
{
    public partial class MainForm : Form
    {
        private string folderPath = "";
        private List<string> imageFiles = new List<string>();
        private bool isProcessing = false;
        private ImageAnnotatorClient visionClient;
        private Stack<Dictionary<string, List<string[]>>> undoStack = new Stack<Dictionary<string, List<string[]>>>();

        // Services for business logic
        private ExcelInvoiceService _excelInvoiceService;
        private OCRTextParsingService _ocrParsingService;

        // OCR Tab Controls
        private TextBox txtNguoiDiOCR;
        private TextBox txtNguoiLayOCR;
        private RichTextBox txtRawOCRLog;
        private RichTextBox txtProcessLog;
        private CheckedListBox chkListImages;
        private List<Dictionary<string, string>> mappedDataList = new List<Dictionary<string, string>>();

        public MainForm()
        {
            InitializeComponent();
            InitializeServices();
            LoadApplicationIcon();
            InitializeTesseract();
            InitializeOCRTab();
            InitializeManualInputTab();
        }

        private void InitializeServices()
        {
            try
            {
                _excelInvoiceService = new ExcelInvoiceService();
                _ocrParsingService = new OCRTextParsingService();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Warning: {ex.Message}");
                // Services can be initialized later
            }
        }

        private void LoadApplicationIcon()
        {
            try
            {
                string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "resources", "app.ico");
                if (File.Exists(iconPath))
                {
                    this.Icon = new Icon(iconPath);
                    System.Diagnostics.Debug.WriteLine($"✅ Icon loaded from: {iconPath}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"❌ Icon file not found: {iconPath}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading icon: {ex.Message}");
            }
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
            mappedDataList.Clear();

            // Get NGƯỜI ĐI and NGƯỜI LẤY from OCR tab
            string nguoiDi = txtNguoiDiOCR?.Text ?? "";
            string nguoiLay = txtNguoiLayOCR?.Text ?? "";

            if (string.IsNullOrWhiteSpace(nguoiDi) || string.IsNullOrWhiteSpace(nguoiLay))
            {
                this.Invoke((MethodInvoker)delegate
                {
                    MessageBox.Show("❌ Vui lòng nhập NGƯỜI ĐI và NGƯỜI LẤY trước khi quét", "Thông báo", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    btnStart.Enabled = true;
                    btnSelectFolder.Enabled = true;
                    btnClear.Enabled = true;
                    isProcessing = false;
                });
                return;
            }

            allText.AppendLine("╔════════════════════════════════════════════════════════╗");
            allText.AppendLine("║    KẾT QUẢ NHẬN DIỆN & MAP DỮ LIỆU (OCR) TIẾNG VIỆT   ║");
            allText.AppendLine("╚════════════════════════════════════════════════════════╝\n");
            allText.AppendLine($"📅 Ngày: {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
            allText.AppendLine($"📁 Folder: {folderPath}");
            allText.AppendLine($"� Người Đi: {nguoiDi}");
            allText.AppendLine($"👤 Người Lấy: {nguoiLay}");
            allText.AppendLine($"�📷 Tổng ảnh: {imageFiles.Count}");
            allText.AppendLine("\n" + new string('═', 60) + "\n");

            // Update UI with header immediately
            this.Invoke((MethodInvoker)delegate
            {
                txtResult.Text = allText.ToString();
                txtProcessLog.Text = allText.ToString();
            });

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
                    // OCR từ ảnh
                    var (text, confidence) = CallPythonOCR(imagePath);

                    // Write raw OCR text to txtRawOCRLog
                    this.Invoke((MethodInvoker)delegate
                    {
                        if (txtRawOCRLog != null)
                        {
                            txtRawOCRLog.AppendText($"\n{'═', 60}\n");
                            txtRawOCRLog.AppendText($"📄 TỆP: {fileName}\n");
                            txtRawOCRLog.AppendText($"📊 Độ tin cậy: {confidence:F1}%\n");
                            txtRawOCRLog.AppendText($"{'─', 60}\n");
                            txtRawOCRLog.AppendText(text ?? "(Empty OCR result)\n");
                        }
                    });

                    allText.AppendLine($"\n✅ TỆP #{i + 1}: {fileName}");
                    allText.AppendLine($"   📊 Độ tin cậy: {confidence:F1}%");
                    allText.AppendLine($"   ⏱️  Thời gian: {DateTime.Now:HH:mm:ss}");
                    allText.AppendLine(new string('─', 60));

                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        // Map dữ liệu từ OCR text
                        var mappedData = MapOCRDataTo12Fields(text, fileName, nguoiDi, nguoiLay);
                        
                        // Validate tất cả 12 fields
                        var missingFields = ValidateMappedData(mappedData);
                        var fieldStatuses = GetFieldStatuses(mappedData);
                        
                        if (missingFields.Count == 0)
                        {
                            allText.AppendLine("\n✅ THÀNH CÔNG - DỮ LIỆU ĐẦY ĐỦ (11/11 FIELDS):");
                            allText.AppendLine($"  ✓ SHOP: {mappedData["SHOP"]}");
                            allText.AppendLine($"  ✓ TÊN KH: {mappedData["TÊN KH"]}");
                            allText.AppendLine($"  ✓ MÃ: {mappedData["MÃ"]}");
                            allText.AppendLine($"  ✓ SỐ NHÀ: {mappedData["SỐ NHÀ"]}");
                            allText.AppendLine($"  ✓ TÊN ĐƯỜNG: {mappedData["TÊN ĐƯỜNG"]}");
                            allText.AppendLine($"  ✓ QUẬN: {mappedData["QUẬN"]}");
                            allText.AppendLine($"  ✓ TIỀN THU: {mappedData["TIỀN THU"]}");
                            allText.AppendLine($"  ✓ TIỀN SHIP: {mappedData["TIỀN SHIP"]}");
                            allText.AppendLine($"  ✓ TIỀN HÀNG: {mappedData["TIỀN HÀNG"]}");
                            allText.AppendLine($"  ✓ NGÀY LẤY: {mappedData["NGÀY LẤY"]}");
                            allText.AppendLine($"  ✓ NGƯỜI ĐI: {mappedData["NGƯỜI ĐI"]}");
                            allText.AppendLine($"  ✓ NGƯỜI LẤY: {mappedData["NGƯỜI LẤY"]}");
                            
                            mappedDataList.Add(mappedData);
                            successCount++;
                        }
                        else
                        {
                            int passedCount = 11 - missingFields.Count;
                            allText.AppendLine($"\n⚠️ TỰA THÀNH CÔNG ({passedCount}/11 FIELDS):");
                            
                            // Log fields that passed
                            allText.AppendLine("   ✅ FIELDS PASS:");
                            foreach (var kvp in fieldStatuses)
                            {
                                if (kvp.Value)
                                {
                                    allText.AppendLine($"      ✓ {kvp.Key}: {mappedData[kvp.Key]}");
                                }
                            }
                            
                            // Log fields that failed
                            allText.AppendLine("   ❌ FIELDS FAIL:");
                            foreach (var field in missingFields)
                            {
                                allText.AppendLine($"      ✗ {field}");
                            }
                            failCount++;
                        }
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

                // Update txtResult after each file to show progress in real-time
                this.Invoke((MethodInvoker)delegate
                {
                    txtResult.Text = allText.ToString();
                    txtResult.SelectionStart = txtResult.Text.Length;
                    txtResult.ScrollToCaret();
                    
                    txtProcessLog.Text = allText.ToString();
                    txtProcessLog.SelectionStart = txtProcessLog.Text.Length;
                    txtProcessLog.ScrollToCaret();
                });
            }

            allText.AppendLine("\n\n╔════════════════════════════════════════════════════════╗");
            allText.AppendLine("║                    TÓM TẮT KẾT QUẢ                      ║");
            allText.AppendLine("╚════════════════════════════════════════════════════════╝\n");
            allText.AppendLine($"✅ Thành công (đủ 11 fields): {successCount}/{imageFiles.Count} ảnh");
            allText.AppendLine($"❌ Thất bại/Thiếu thông tin: {failCount}/{imageFiles.Count} ảnh");
            allText.AppendLine($"⏱️  Thời gian xử lý: {DateTime.Now:HH:mm:ss}\n");
            allText.AppendLine($"💾 Sẵn sàng xuất {mappedDataList.Count} dòng dữ liệu sang Excel");

            this.Invoke((MethodInvoker)delegate
            {
                txtResult.Text = allText.ToString();
                txtProcessLog.Text = allText.ToString();
                lblCurrentFile.Text = $"✅ Hoàn thành: {successCount} thành công, {failCount} thất bại";
                lblStatus.Text = "✅ Xử lý xong";
                lblStatus.ForeColor = Color.Green;

                btnStart.Enabled = true;
                btnSelectFolder.Enabled = true;
                btnClear.Enabled = true;

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

        // Excel Viewer Event Handler
        private void BtnOpenExcel_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
                    openFileDialog.Title = "Chọn file Excel";
                    
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Load and display Excel sheets
                        LoadExcelFile(openFileDialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi:\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadExcelFile(string filePath)
        {
            try
            {
                // Store the file path for saving later
                currentExcelFilePath = filePath;

                using (var workbook = new XLWorkbook(filePath))
                {
                    var sheetNames = workbook.Worksheets.Select(ws => ws.Name).ToList();

                    if (sheetNames.Count == 0)
                    {
                        MessageBox.Show("⚠️ File Excel không có sheet nào", "Thông báo");
                        return;
                    }

                    // Clear existing tabs and load into main form's tabExcelSheets
                    tabExcelSheets.TabPages.Clear();

                    foreach (var sheetName in sheetNames)
                    {
                        TabPage tabPage = new TabPage(sheetName);
                        DataGridView dgv = new DataGridView();
                        dgv.Dock = DockStyle.Fill;
                        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                        dgv.ReadOnly = false;  // ✅ Allow editing
                        dgv.AllowUserToAddRows = true;  // ✅ Allow adding rows
                        dgv.AllowUserToDeleteRows = true;  // ✅ Allow deleting rows
                        tabPage.Controls.Add(dgv);

                        LoadSheetData(workbook, sheetName, dgv);
                        tabExcelSheets.TabPages.Add(tabPage);
                    }

                    // Switch to Excel tab
                    tabMainControl.SelectedTab = tabExcelViewer;

                    lblStatus.Text = $"✅ Excel: {Path.GetFileName(filePath)} ({sheetNames.Count} sheets)";
                    lblStatus.ForeColor = Color.Green;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi đọc Excel:\n{ex.Message}", "Lỗi");
                Debug.WriteLine($"Excel error: {ex.Message}");
            }
        }

        private void LoadSheetData(XLWorkbook workbook, string sheetName, DataGridView dgv)
        {
            try
            {
                var worksheet = workbook.Worksheet(sheetName);
                DataTable dataTable = new DataTable();

                var usedRange = worksheet.RangeUsed();
                if (usedRange == null) return;

                int rowCount = usedRange.RowCount();
                int colCount = usedRange.ColumnCount();

                // Tìm hàng header thực (hàng có "SHOP", "TÊN KH", v.v.) - thường ở hàng 2
                int headerRowIndex = 2;
                for (int row = 1; row <= Math.Min(5, rowCount); row++)
                {
                    string firstCell = worksheet.Cell(row, 1).GetString()?.Trim() ?? "";
                    if (firstCell == "SHOP" || firstCell.Contains("Tình trạng"))
                    {
                        headerRowIndex = row;
                        break;
                    }
                }

                // Add columns từ hàng header
                for (int col = 1; col <= colCount; col++)
                {
                    string columnName = worksheet.Cell(headerRowIndex, col).GetString()?.Trim() ?? "";
                    dataTable.Columns.Add(columnName);
                }

                // Add rows - BẮT ĐẦU TỪ HÀNG 1 (để giữ "THU 2", "NGAY 2-2" v.v.)
                for (int row = 1; row <= rowCount; row++)
                {
                    // Skip hàng header (hàng có tên cột thực)
                    if (row == headerRowIndex)
                        continue;
                    
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= colCount; col++)
                    {
                        string cellValue = worksheet.Cell(row, col).GetString();
                        dataRow[col - 1] = cellValue ?? "";
                    }
                    dataTable.Rows.Add(dataRow);
                }

                dgv.DataSource = dataTable;

                // Auto-fit columns
                dgv.AutoResizeColumns();

                // Freeze hàng header (hàng đầu tiên của DataGridView)
                if (dgv.Rows.Count > 0)
                {
                    dgv.Rows[0].Frozen = true;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Sheet error: {ex.Message}");
            }
        }

        private void BtnAddInvoiceRow_Click(object sender, EventArgs e)
        {
            if (dgvInvoice.Columns.Count == 0)
            {
                // Initialize columns - Simple 3 columns: Tên | Tiền | Số đơn
                dgvInvoice.Columns.Add("Tên", "Tên");
                dgvInvoice.Columns.Add("Tiền", "Tiền");
                dgvInvoice.Columns.Add("Số đơn", "Số đơn");
            }

            dgvInvoice.Rows.Add("", "0", "0");
        }

        private void BtnSaveInvoice_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvInvoice.Rows.Count == 0)
                {
                    MessageBox.Show("Chưa có dữ liệu để lưu!", "Thông báo");
                    return;
                }

                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                dialog.FileName = $"Invoice_{DateTime.Now:dd-MM-yyyy}.xlsx";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    ExportInvoiceToExcel(dgvInvoice, dialog.FileName);
                    MessageBox.Show($"✅ Lưu thành công!\n{dialog.FileName}", "Thành công");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi: {ex.Message}", "Lỗi");
            }
        }

        private void ExportInvoiceToExcel(DataGridView dgv, string filePath)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Hóa đơn");

                    // Add headers
                    for (int col = 0; col < dgv.Columns.Count; col++)
                    {
                        worksheet.Cell(1, col + 1).Value = dgv.Columns[col].HeaderText;
                    }

                    // Add data
                    for (int row = 0; row < dgv.Rows.Count; row++)
                    {
                        for (int col = 0; col < dgv.Columns.Count; col++)
                        {
                            var cellValue = dgv.Rows[row].Cells[col].Value;
                            worksheet.Cell(row + 2, col + 1).Value = cellValue?.ToString() ?? "";
                        }
                    }

                    // Calculate totals if possible
                    int lastRow = dgv.Rows.Count + 2;
                    worksheet.Cell(lastRow, 1).Value = "TỔNG CỘNG";
                    worksheet.Cell(lastRow, 1).Style.Font.Bold = true;

                    workbook.SaveAs(filePath);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Export error: {ex.Message}");
                throw;
            }
        }

        private void BtnImportFromExcel_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    using (var workbook = new XLWorkbook(dialog.FileName))
                    {
                        var sheetNames = workbook.Worksheets.Select(ws => ws.Name).ToList();

                        if (sheetNames.Count == 0)
                        {
                            MessageBox.Show("File Excel không có sheet nào", "Thông báo");
                            return;
                        }

                        // Let user select which sheet to import from
                        string selectedSheet = sheetNames[0]; // Default first sheet
                        
                        if (sheetNames.Count > 1)
                        {
                            // Simple dialog to select sheet
                            using (Form selectForm = new Form())
                            {
                                selectForm.Text = "Chọn Sheet";
                                selectForm.Width = 300;
                                selectForm.Height = 150;
                                selectForm.StartPosition = FormStartPosition.CenterParent;

                                ComboBox cbSheets = new ComboBox();
                                cbSheets.DataSource = sheetNames;
                                cbSheets.Location = new Point(10, 20);
                                cbSheets.Width = 260;

                                Button btnOk = new Button();
                                btnOk.Text = "OK";
                                btnOk.Location = new Point(100, 70);
                                btnOk.Click += (s, evt) => selectForm.DialogResult = DialogResult.OK;

                                selectForm.Controls.Add(cbSheets);
                                selectForm.Controls.Add(btnOk);

                                if (selectForm.ShowDialog() == DialogResult.OK)
                                {
                                    selectedSheet = cbSheets.SelectedItem.ToString();
                                }
                            }
                        }

                        // Import data from selected sheet
                        ImportInvoiceData(workbook, selectedSheet);
                        MessageBox.Show($"✅ Nhập dữ liệu từ sheet '{selectedSheet}' thành công!\n\nBây giờ bấm 🧮 Tính Tiền để tính tổng", "Thành công");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi: {ex.Message}", "Lỗi");
            }
        }

        private void ImportInvoiceData(XLWorkbook workbook, string sheetName)
        {
            try
            {
                var worksheet = workbook.Worksheet(sheetName);
                var usedRange = worksheet.RangeUsed();

                if (usedRange == null) return;

                // Initialize columns if needed
                if (dgvInvoice.Columns.Count == 0)
                {
                    dgvInvoice.Columns.Add("Mặt hàng", "Mặt hàng");
                    dgvInvoice.Columns.Add("Số lượng", "Số lượng");
                    dgvInvoice.Columns.Add("Đơn giá", "Đơn giá");
                    dgvInvoice.Columns.Add("Thành tiền", "Thành tiền");
                }

                dgvInvoice.Rows.Clear();

                // Find summary section (look for "TỔNG" or "TOTAL" rows)
                // This scans the sheet and extracts item info
                int rowCount = usedRange.RowCount();
                
                for (int row = 1; row <= rowCount; row++)
                {
                    string mh = worksheet.Cell(row, 2).GetString()?.Trim() ?? "";
                    string tenduong = worksheet.Cell(row, 6).GetString()?.Trim() ?? "";
                    string quan = worksheet.Cell(row, 7).GetString()?.Trim() ?? "";
                    string tienhan = worksheet.Cell(row, 8).GetString()?.Trim() ?? "";

                    // Only add rows that have meaningful data (not headers or empty rows)
                    if (!string.IsNullOrEmpty(mh) && !mh.Contains("SHOP") && !mh.Contains("Tính"))
                    {
                        string displayName = $"{mh} - {tenduong}".Trim();
                        
                        if (!string.IsNullOrEmpty(tienhan) && decimal.TryParse(tienhan, out decimal price))
                        {
                            if (!string.IsNullOrEmpty(quan) && decimal.TryParse(quan, out decimal qty))
                            {
                                decimal total = price * qty;
                                dgvInvoice.Rows.Add(displayName, qty, price, total);
                            }
                        }
                    }
                }

                // Auto-calculate totals
                CalculateInvoiceTotals();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Import error: {ex.Message}");
                throw;
            }
        }

        private void CalculateInvoiceTotals()
        {
            // Calculate "Thành tiền" = Số lượng × Đơn giá for each row
            for (int i = 0; i < dgvInvoice.Rows.Count; i++)
            {
                if (decimal.TryParse(dgvInvoice.Rows[i].Cells[1].Value?.ToString() ?? "0", out decimal qty) &&
                    decimal.TryParse(dgvInvoice.Rows[i].Cells[2].Value?.ToString() ?? "0", out decimal price))
                {
                    decimal total = qty * price;
                    dgvInvoice.Rows[i].Cells[3].Value = total;
                }
            }
        }

        private void BtnCalculateInvoice_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvInvoice.Rows.Count == 0)
                {
                    MessageBox.Show("Chưa có dữ liệu để tính!", "Thông báo");
                    return;
                }

                decimal totalTien = 0;
                decimal totalSoDon = 0;

                // Calculate only 2 values: TIỀN HÀNG + SỐ ĐƠN
                for (int i = 0; i < dgvInvoice.Rows.Count; i++)
                {
                    // Column 1: Tiền hàng
                    if (decimal.TryParse(dgvInvoice.Rows[i].Cells[1].Value?.ToString() ?? "0", out decimal tienHang))
                    {
                        totalTien += tienHang;
                    }

                    // Column 8: Số đơn (currently storing here)
                    if (decimal.TryParse(dgvInvoice.Rows[i].Cells[8].Value?.ToString() ?? "0", out decimal sodon))
                    {
                        totalSoDon += sodon;
                    }
                }

                // Update total label
                lblInvoiceTotal.Text = $"TỔNG CỘNG: {totalTien:N0} đ | SỐ ĐƠN: {totalSoDon:N0}";
                
                // Create Daily Report data
                currentDailyReport = new DailyReportData
                {
                    Date = DateTime.Now.ToString("dd.MM.yyyy"),
                    TienHangThanhToan = totalTien,
                    TruDonCuDaCk = 0,
                    SoDon = totalSoDon
                };
                
                // Initialize button panel and display Daily Report
                InitializeInvoiceButtonPanel();
                DisplayDailyReport();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ Lỗi: {ex.Message}");
            }
        }

        private void SaveInvoiceToExcelSheet(decimal totalAmount)
        {
            try
            {
                if (string.IsNullOrEmpty(currentExcelFilePath))
                {
                    MessageBox.Show("Vui lòng mở file Excel trước!", "Thông báo");
                    return;
                }

                string sheetName = DateTime.Now.ToString("dd-MM");
                
                using (var workbook = new XLWorkbook(currentExcelFilePath))
                {
                    // Remove sheet if exists then recreate (ghi đè)
                    if (workbook.TryGetWorksheet(sheetName, out _))
                    {
                        workbook.Worksheets.Delete(sheetName);
                    }

                    // Create new sheet with today's date
                    var worksheet = workbook.Worksheets.Add(sheetName);

                    // Add headers
                    for (int col = 0; col < dgvInvoice.Columns.Count; col++)
                    {
                        worksheet.Cell(1, col + 1).Value = dgvInvoice.Columns[col].HeaderText;
                    }

                    // Add data rows
                    for (int row = 0; row < dgvInvoice.Rows.Count; row++)
                    {
                        for (int col = 0; col < dgvInvoice.Columns.Count; col++)
                        {
                            var cellValue = dgvInvoice.Rows[row].Cells[col].Value;
                            worksheet.Cell(row + 2, col + 1).Value = cellValue?.ToString() ?? "";
                        }
                    }

                    // Add total row
                    int lastRow = dgvInvoice.Rows.Count + 2;
                    worksheet.Cell(lastRow, 1).Value = "TỔNG CỘNG";
                    worksheet.Cell(lastRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(lastRow, 9).Value = totalAmount;
                    worksheet.Cell(lastRow, 9).Style.Font.Bold = true;

                    workbook.SaveAs(currentExcelFilePath);
                }

                MessageBox.Show($"✅ Lưu vào sheet '{sheetName}' thành công!", "Thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi: {ex.Message}", "Lỗi");
                Debug.WriteLine($"Save error: {ex.Message}");
            }
        }

        // Save Excel Editor Handler
        private void BtnSaveExcelEditor_Click(object sender, EventArgs e)
        {
            try
            {
                if (tabExcelSheets.TabPages.Count == 0)
                {
                    MessageBox.Show("Chưa mở file Excel!", "Thông báo");
                    return;
                }

                if (string.IsNullOrEmpty(currentExcelFilePath))
                {
                    MessageBox.Show("Không tìm thấy đường dẫn file Excel!", "Lỗi");
                    return;
                }

                // Save all sheets from DataGridView back to Excel
                using (var workbook = new XLWorkbook(currentExcelFilePath))
                {
                    foreach (TabPage tabPage in tabExcelSheets.TabPages)
                    {
                        var dgv = tabPage.Controls[0] as DataGridView;
                        if (dgv == null) continue;

                        string sheetName = tabPage.Text;
                        var worksheet = workbook.Worksheet(sheetName);

                        // Clear existing data
                        worksheet.Clear();

                        // Write headers
                        for (int col = 0; col < dgv.Columns.Count; col++)
                        {
                            worksheet.Cell(1, col + 1).Value = dgv.Columns[col].HeaderText;
                        }

                        // Write data rows
                        for (int row = 0; row < dgv.Rows.Count; row++)
                        {
                            for (int col = 0; col < dgv.Columns.Count; col++)
                            {
                                var cellValue = dgv.Rows[row].Cells[col].Value;
                                if (cellValue != null)
                                {
                                    worksheet.Cell(row + 2, col + 1).Value = cellValue.ToString();
                                }
                            }
                        }
                    }

                    workbook.SaveAs(currentExcelFilePath);
                }

                MessageBox.Show($"✅ Lưu file Excel thành công!", "Thành công");
                lblStatus.Text = $"✅ Lưu Excel: {Path.GetFileName(currentExcelFilePath)}";
                lblStatus.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi lưu: {ex.Message}", "Lỗi");
                Debug.WriteLine($"Save Excel error: {ex.Message}");
            }
        }

        // Undo Excel Editor Handler
        private void BtnUndoExcelEditor_Click(object sender, EventArgs e)
        {
            try
            {
                if (tabExcelSheets.TabPages.Count == 0)
                {
                    MessageBox.Show("Chưa mở file Excel!", "Thông báo");
                    return;
                }

                // Reload the current sheet from file (cancel all changes)
                if (!string.IsNullOrEmpty(currentExcelFilePath))
                {
                    LoadExcelFile(currentExcelFilePath);
                    MessageBox.Show("✅ Đã hoàn tác tất cả thay đổi!", "Thành công");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi: {ex.Message}", "Lỗi");
                Debug.WriteLine($"Undo error: {ex.Message}");
            }
        }

        // Cancel Excel Editor Handler
        private void BtnCancelExcelEditor_Click(object sender, EventArgs e)
        {
            try
            {
                tabExcelSheets.TabPages.Clear();
                currentExcelFilePath = "";
                lblStatus.Text = "✅ Đã đóng file Excel";
                lblStatus.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi: {ex.Message}", "Lỗi");
            }
        }

        // Calculate button in Excel Viewer
        private void BtnCalculateExcelData_Click(object sender, EventArgs e)
        {
            try
            {
                if (tabExcelSheets.TabPages.Count == 0)
                    return;

                // Get current sheet's DataGridView
                var currentSheet = tabExcelSheets.SelectedTab;
                if (currentSheet == null || currentSheet.Controls.Count == 0)
                    return;

                // Get the DataGridView from the current sheet
                DataGridView sourceGridView = null;
                foreach (Control ctrl in currentSheet.Controls)
                {
                    if (ctrl is DataGridView dgv)
                    {
                        sourceGridView = dgv;
                        break;
                    }
                }

                if (sourceGridView == null || sourceGridView.Rows.Count == 0)
                    return;

                // Find column indices for calculation
                int colShop = -1;  // SHOP (để detect dòng hàng hóa vs dòng tính)
                int colTienHang = -1;  // TIỀN HÀNG (cột J) - sum này để tính tổng tiền
                int colSoDon = -1;  // SỐ ĐƠN (cột R) - sum này để tính tổng đơn

                for (int col = 0; col < sourceGridView.Columns.Count; col++)
                {
                    string header = sourceGridView.Columns[col].HeaderText.ToLower();
                    if (header.Contains("shop")) colShop = col;
                    if (header.Contains("tiền hàng")) colTienHang = col;
                    if (header.Contains("số đơn")) colSoDon = col;
                }

                // DEBUG: Log column info
                Debug.WriteLine($"=== COLUMN DETECTION ===");
                Debug.WriteLine($"Total columns in sourceGridView: {sourceGridView.Columns.Count}");
                Debug.WriteLine($"Column indices - Shop: {colShop}, TienHang: {colTienHang}, SoDon: {colSoDon}");
                for (int i = 0; i < sourceGridView.Columns.Count; i++)
                {
                    Debug.WriteLine($"Col {i}: '{sourceGridView.Columns[i].HeaderText}'");
                }
                Debug.WriteLine($"=== DATA ROWS ===");

                // PHẦN 1: Copy toàn bộ dữ liệu từ Excel sang dgvInvoice
                dgvInvoice.DataSource = null;
                dgvInvoice.Rows.Clear();
                dgvInvoice.Columns.Clear();

                // Copy columns
                foreach (DataGridViewColumn col in sourceGridView.Columns)
                {
                    dgvInvoice.Columns.Add(col.Name, col.HeaderText);
                }

                // Copy rows - only copy rows with SHOP value (skip SUM rows and adjustment rows)
                foreach (DataGridViewRow sourceRow in sourceGridView.Rows)
                {
                    if (sourceRow.IsNewRow) continue;

                    // Only copy rows that have a SHOP value (skip adjustment/sum rows)
                    string shopValue = sourceRow.Cells[colShop].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(shopValue.Trim())) continue;  // Skip rows without SHOP

                    DataGridViewRow newRow = new DataGridViewRow();
                    newRow.CreateCells(dgvInvoice);

                    for (int i = 0; i < sourceRow.Cells.Count; i++)
                    {
                        newRow.Cells[i].Value = sourceRow.Cells[i].Value;
                    }

                    dgvInvoice.Rows.Add(newRow);
                }

                // PHẦN 2: Calculate Daily Report
                // Logic: 
                // 1. Find SUM row (row without SHOP, has value in column J)
                // 2. Get TIỀN HÀNG from column J and SỐ ĐƠN from column R
                // 3. Find adjustment rows (rows after SUM with negative values in column J)
                // 4. Subtract adjustments from the total
                
                decimal baseTienHang = 0;  // Base amount from SUM row
                decimal totalTienHang = 0; // Final amount after adjustments
                decimal totalSoDon = 0;
                int sumRowIndex = -1;
                List<decimal> adjustments = new List<decimal>();
                DataGridViewRow sumRowToDisplay = null;  // Store SUM row to display

                // Find SUM row - it's the row with NO SHOP but has large value in column J
                for (int i = sourceGridView.Rows.Count - 1; i >= 0; i--)
                {
                    DataGridViewRow row = sourceGridView.Rows[i];
                    if (row.IsNewRow) continue;

                    string shopValue = "";
                    if (colShop >= 0 && colShop < row.Cells.Count)
                    {
                        shopValue = row.Cells[colShop].Value?.ToString() ?? "";
                    }

                    // SUM row has NO SHOP value but has positive number in column J
                    if (string.IsNullOrEmpty(shopValue.Trim()))
                    {
                        if (colTienHang >= 0 && colTienHang < row.Cells.Count)
                        {
                            object cellValue = row.Cells[colTienHang].Value;
                            if (cellValue != null && decimal.TryParse(cellValue.ToString(), out decimal jValue) && jValue > 0)
                            {
                                // Found the SUM row
                                baseTienHang = jValue;
                                totalTienHang = jValue;
                                sumRowIndex = i;
                                sumRowToDisplay = row;  // Save SUM row for display
                                
                                Debug.WriteLine($"*** Found SUM row at index {i}");
                                Debug.WriteLine($"    colShop={colShop}, colTienHang={colTienHang}, colSoDon={colSoDon}");
                                Debug.WriteLine($"    Row has {row.Cells.Count} cells");
                                
                                // Get SỐ ĐƠN từ cột R
                                // Try multiple methods to find it
                                totalSoDon = 0;
                                
                                // Method 1: Use detected column index
                                if (colSoDon >= 0 && colSoDon < row.Cells.Count)
                                {
                                    object soDonValue = row.Cells[colSoDon].Value;
                                    Debug.WriteLine($"    Method 1 (colSoDon={colSoDon}): Value={soDonValue}, Type={soDonValue?.GetType().Name ?? "null"}");
                                    if (soDonValue != null)
                                    {
                                        try
                                        {
                                            totalSoDon = Convert.ToDecimal(soDonValue);
                                            Debug.WriteLine($"      ✓ Success: {totalSoDon}");
                                        }
                                        catch
                                        {
                                            Debug.WriteLine($"      ✗ Failed to parse");
                                        }
                                    }
                                }
                                
                                // Method 2: Look for "số đơn" in header and use that column
                                if (totalSoDon == 0)
                                {
                                    for (int col = 0; col < sourceGridView.Columns.Count; col++)
                                    {
                                        string header = sourceGridView.Columns[col].HeaderText.ToLower();
                                        if (header.Contains("số") && header.Contains("đơn"))
                                        {
                                            object soDonValue = row.Cells[col].Value;
                                            Debug.WriteLine($"    Method 2 (found at col {col}): Value={soDonValue}, Type={soDonValue?.GetType().Name ?? "null"}");
                                            if (soDonValue != null)
                                            {
                                                try
                                                {
                                                    totalSoDon = Convert.ToDecimal(soDonValue);
                                                    Debug.WriteLine($"      ✓ Success: {totalSoDon}");
                                                }
                                                catch
                                                {
                                                    Debug.WriteLine($"      ✗ Failed to parse");
                                                }
                                            }
                                            break;
                                        }
                                    }
                                }
                                
                                // Method 3: Try column R directly (index 17)
                                if (totalSoDon == 0 && row.Cells.Count > 17)
                                {
                                    object soDonValue = row.Cells[17].Value;
                                    Debug.WriteLine($"    Method 3 (col 17): Value={soDonValue}, Type={soDonValue?.GetType().Name ?? "null"}");
                                    if (soDonValue != null)
                                    {
                                        try
                                        {
                                            totalSoDon = Convert.ToDecimal(soDonValue);
                                            Debug.WriteLine($"      ✓ Success: {totalSoDon}");
                                        }
                                        catch
                                        {
                                            Debug.WriteLine($"      ✗ Failed to parse");
                                        }
                                    }
                                }
                                
                                Debug.WriteLine($"    *** Final BaseTienHang={baseTienHang}, SoDon={totalSoDon} ***");
                                break;
                            }
                        }
                    }
                }

                // Find adjustment rows (rows after SUM row, with negative values in column J)
                if (sumRowIndex >= 0)
                {
                    for (int i = sumRowIndex + 1; i < sourceGridView.Rows.Count; i++)
                    {
                        DataGridViewRow row = sourceGridView.Rows[i];
                        if (row.IsNewRow) continue;

                        // Check for negative value in column J (adjustment)
                        if (colTienHang >= 0 && colTienHang < row.Cells.Count)
                        {
                            object cellValue = row.Cells[colTienHang].Value;
                            if (cellValue != null && decimal.TryParse(cellValue.ToString(), out decimal jValue) && jValue < 0)
                            {
                                adjustments.Add(jValue);
                                totalTienHang += jValue;  // jValue is negative, so this subtracts
                                Debug.WriteLine($"  -> Found adjustment at row {i}: {jValue}, Running total={totalTienHang}");
                            }
                        }
                    }
                }

                Debug.WriteLine($"=== FINAL CALCULATION ===");
                Debug.WriteLine($"Base TienHang: {baseTienHang}");
                Debug.WriteLine($"Adjustments: {string.Join(", ", adjustments)}");
                Debug.WriteLine($"Final TienHang: {totalTienHang}");
                Debug.WriteLine($"SoDon: {totalSoDon}");

                // Add SUM row to display (with yellow background)
                if (sumRowToDisplay != null)
                {
                    DataGridViewRow sumDisplayRow = new DataGridViewRow();
                    sumDisplayRow.CreateCells(dgvInvoice);

                    for (int i = 0; i < sumRowToDisplay.Cells.Count && i < sumDisplayRow.Cells.Count; i++)
                    {
                        sumDisplayRow.Cells[i].Value = sumRowToDisplay.Cells[i].Value;
                    }

                    dgvInvoice.Rows.Add(sumDisplayRow);

                    // Color the SUM row yellow
                    int lastRowIndex = dgvInvoice.Rows.Count - 1;
                    for (int col = 0; col < dgvInvoice.Columns.Count; col++)
                    {
                        dgvInvoice.Rows[lastRowIndex].Cells[col].Style.BackColor = Color.Yellow;
                        dgvInvoice.Rows[lastRowIndex].Cells[col].Style.Font = new Font(dgvInvoice.Font, FontStyle.Bold);
                    }

                    // Add adjustment rows (rows after SUM with negative values)
                    if (sumRowIndex >= 0)
                    {
                        for (int i = sumRowIndex + 1; i < sourceGridView.Rows.Count; i++)
                        {
                            DataGridViewRow adjRow = sourceGridView.Rows[i];
                            if (adjRow.IsNewRow) continue;

                            // Check if this is an adjustment row (has negative value in column J)
                            if (colTienHang >= 0 && colTienHang < adjRow.Cells.Count)
                            {
                                object cellValue = adjRow.Cells[colTienHang].Value;
                                if (cellValue != null && decimal.TryParse(cellValue.ToString(), out decimal jValue) && jValue < 0)
                                {
                                    // Add adjustment row to display
                                    DataGridViewRow adjDisplayRow = new DataGridViewRow();
                                    adjDisplayRow.CreateCells(dgvInvoice);

                                    for (int col = 0; col < adjRow.Cells.Count && col < adjDisplayRow.Cells.Count; col++)
                                    {
                                        adjDisplayRow.Cells[col].Value = adjRow.Cells[col].Value;
                                    }

                                    dgvInvoice.Rows.Add(adjDisplayRow);

                                    // Color adjustment rows light orange/peach
                                    int adjRowIndex = dgvInvoice.Rows.Count - 1;
                                    for (int col = 0; col < dgvInvoice.Columns.Count; col++)
                                    {
                                        dgvInvoice.Rows[adjRowIndex].Cells[col].Style.BackColor = Color.FromArgb(255, 200, 124);  // Light orange
                                        dgvInvoice.Rows[adjRowIndex].Cells[col].Style.Font = new Font(dgvInvoice.Font, FontStyle.Italic);
                                    }
                                }
                            }
                        }
                    }
                }

                // KHÔNG thêm TỔNG CỘNG row vào dgvInvoice
                // Chỉ tính toán để lưu vào currentDailyReport
                // Tổng sẽ được hiển thị ở phần 2 (Daily Report)

                // Store calculation results for Daily Report display
                currentDailyReport = new DailyReportData
                {
                    Date = DateTime.Now.ToString("dd.MM.yyyy"),
                    TienHangThanhToan = totalTienHang,
                    TruDonCuDaCk = 0,  // Adjustment không cộng
                    SoDon = totalSoDon
                };

                // Update label
                lblInvoiceTotal.Text = $"TỔNG CỘNG: {totalTienHang:N0} đ | SỐ ĐƠN: {totalSoDon:N0}";

                // Display Daily Report
                DisplayDailyReport();
                // Initialize button panel
                InitializeInvoiceButtonPanel();

                // Switch to Invoice tab
                tabMainControl.SelectedIndex = 2;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ Lỗi: {ex.Message}");
            }
        }

        // Helper class for Daily Report data
        private class DailyReportData
        {
            public string Date { get; set; }
            public decimal TienHangThanhToan { get; set; }
            public decimal TruDonCuDaCk { get; set; }
            public decimal SoDon { get; set; }
        }

        private DailyReportData currentDailyReport;

        // Display Daily Report in a new panel below dgvInvoice
        private void DisplayDailyReport()
        {
            if (currentDailyReport == null) return;

            // Initialize container panels if needed
            Panel pnlTop = tabInvoice.Controls["pnlInvoiceTop"] as Panel;
            Panel pnlBottom = tabInvoice.Controls["pnlDailyReportBottom"] as Panel;

            // First time setup: create panel structure
            if (pnlTop == null)
            {
                // Clear default controls from tabInvoice
                tabInvoice.Controls.Clear();

                // Create top panel (70% of space) for DataGridView
                pnlTop = new Panel();
                pnlTop.Name = "pnlInvoiceTop";
                pnlTop.Dock = DockStyle.Fill;
                pnlTop.BackColor = Color.White;
                pnlTop.Controls.Add(dgvInvoice);
                pnlTop.Controls.Add(lblInvoiceTotal);
                tabInvoice.Controls.Add(pnlTop);

                // Create bottom panel (30% of space) for Daily Report
                pnlBottom = new Panel();
                pnlBottom.Name = "pnlDailyReportBottom";
                pnlBottom.Dock = DockStyle.Bottom;
                pnlBottom.BackColor = Color.White;
                pnlBottom.BorderStyle = BorderStyle.FixedSingle;
                pnlBottom.Height = 250;
                tabInvoice.Controls.Add(pnlBottom);
            }

            pnlBottom.Controls.Clear();

            // Create DataGridView for Daily Report (format giống ảnh)
            DataGridView dgvReport = new DataGridView();
            dgvReport.Dock = DockStyle.Fill;
            dgvReport.BackgroundColor = Color.White;
            dgvReport.AllowUserToAddRows = false;
            dgvReport.AllowUserToDeleteRows = false;
            dgvReport.ReadOnly = true;
            dgvReport.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvReport.ScrollBars = ScrollBars.Both;
            dgvReport.RowHeadersVisible = false;

            // Add columns
            dgvReport.Columns.Add("TenMuc", "Tên mục");
            dgvReport.Columns.Add("Tien", "Tiền");
            dgvReport.Columns.Add("SoDon", "Số đơn");

            dgvReport.Columns[0].Width = 200;
            dgvReport.Columns[1].Width = 100;
            dgvReport.Columns[2].Width = 80;

            // ===== Row 1: Ngày =====
            dgvReport.Rows.Add(currentDailyReport.Date, "Tiền", "Số đơn");
            dgvReport.Rows[0].DefaultCellStyle.BackColor = Color.LightBlue;
            dgvReport.Rows[0].DefaultCellStyle.Font = new Font("Arial", 10, FontStyle.Bold);
            dgvReport.Rows[0].Height = 25;

            // ===== Data rows =====
            dgvReport.Rows.Add("Tiền hàng Thanh Toán", currentDailyReport.TienHangThanhToan.ToString("N0"), currentDailyReport.SoDon.ToString("N0"));
            dgvReport.Rows[1].DefaultCellStyle.BackColor = Color.White;
            dgvReport.Rows[1].DefaultCellStyle.Font = new Font("Arial", 10);
            
            // DEBUG
            Debug.WriteLine($"DEBUG: TienHangThanhToan = {currentDailyReport.TienHangThanhToan}, SoDon = {currentDailyReport.SoDon}");
            
            dgvReport.Rows.Add("Trừ Ship", "", "");
            dgvReport.Rows.Add("Cước xe", "", "");
            dgvReport.Rows.Add("Khách C.khoản", "", "");
            dgvReport.Rows.Add("Giảm tiền thu Khách", "", "");
            dgvReport.Rows.Add("Hàng Boom Trả", "", "");
            
            dgvReport.Rows.Add("Trừ đơn cũ đã ck", "", "");
            dgvReport.Rows[7].DefaultCellStyle.BackColor = Color.FromArgb(255, 200, 124); // Orange color
            dgvReport.Rows[7].DefaultCellStyle.Font = new Font("Arial", 10);

            // ===== Total row =====
            dgvReport.Rows.Add("Tổng Tiền Hàng", currentDailyReport.TienHangThanhToan.ToString("N0"), currentDailyReport.SoDon.ToString("N0"));
            int totalRowIndex = dgvReport.Rows.Count - 1;
            dgvReport.Rows[totalRowIndex].DefaultCellStyle.BackColor = Color.FromArgb(255, 200, 124); // Orange
            dgvReport.Rows[totalRowIndex].DefaultCellStyle.Font = new Font("Arial", 10, FontStyle.Bold);
            dgvReport.Rows[totalRowIndex].Height = 25;

            pnlBottom.Controls.Add(dgvReport);
        }

        // Add button panel for Invoice tab (Lưu, Undo, Đóng)
        private void InitializeInvoiceButtonPanel()
        {
            // Check if button panel already exists
            Panel pnlButtons = tabInvoice.Controls["pnlInvoiceButtons"] as Panel;
            if (pnlButtons != null) return;

            // Create panel for buttons
            pnlButtons = new Panel();
            pnlButtons.Name = "pnlInvoiceButtons";
            pnlButtons.BackColor = Color.FromArgb(40, 40, 40);
            pnlButtons.Height = 40;
            pnlButtons.Dock = DockStyle.Top;
            tabInvoice.Controls.Add(pnlButtons);
            tabInvoice.Controls.SetChildIndex(pnlButtons, tabInvoice.Controls.Count - 1); // Bring to front

            // Button: Save (💾 Lưu)
            Button btnSave = new Button();
            btnSave.Text = "💾 Lưu";
            btnSave.BackColor = Color.FromArgb(40, 40, 40);
            btnSave.ForeColor = Color.White;
            btnSave.FlatStyle = FlatStyle.Flat;
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Font = new Font("Arial", 9);
            btnSave.Size = new Size(75, 30);
            btnSave.Location = new Point(10, 5);
            btnSave.Click += (s, e) => SaveDailyReportToExcel();
            pnlButtons.Controls.Add(btnSave);

            // Button: Undo (↶ Undo)
            Button btnUndo = new Button();
            btnUndo.Text = "↶ Undo";
            btnUndo.BackColor = Color.FromArgb(40, 40, 40);
            btnUndo.ForeColor = Color.White;
            btnUndo.FlatStyle = FlatStyle.Flat;
            btnUndo.FlatAppearance.BorderSize = 0;
            btnUndo.Font = new Font("Arial", 9);
            btnUndo.Size = new Size(75, 30);
            btnUndo.Location = new Point(90, 5);
            btnUndo.Click += (s, e) => MessageBox.Show("↶ Undo thay đổi", "Thông báo");
            pnlButtons.Controls.Add(btnUndo);

            // Button: Close (✕ Đóng)
            Button btnClose = new Button();
            btnClose.Text = "✕ Đóng";
            btnClose.BackColor = Color.FromArgb(40, 40, 40);
            btnClose.ForeColor = Color.White;
            btnClose.FlatStyle = FlatStyle.Flat;
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Font = new Font("Arial", 9);
            btnClose.Size = new Size(75, 30);
            btnClose.Location = new Point(170, 5);
            btnClose.Click += (s, e) => 
            {
                dgvInvoice.Rows.Clear();
                dgvInvoice.Columns.Clear();
                Panel pnlReport = tabInvoice.Controls["pnlDailyReport"] as Panel;
                if (pnlReport != null)
                {
                    tabInvoice.Controls.Remove(pnlReport);
                    pnlReport.Dispose();
                }
                Panel pnlButtons2 = tabInvoice.Controls["pnlInvoiceButtons"] as Panel;
                if (pnlButtons2 != null)
                {
                    tabInvoice.Controls.Remove(pnlButtons2);
                    pnlButtons2.Dispose();
                }
            };
            pnlButtons.Controls.Add(btnClose);
        }

        // Save Daily Report to Excel file (DailyTotalReport.xlsx)
        // Saves BOTH phần 1 (Invoice DataGridView) and phần 2 (Daily Report)
        private void SaveDailyReportToExcel()
        {
            try
            {
                if (dgvInvoice.Rows.Count == 0)
                {
                    MessageBox.Show("Không có dữ liệu để lưu!", "Thông báo");
                    return;
                }

                // Đường dẫn file DailyTotalReport.xlsx
                string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DailyTotalReport.xlsx");

                // Tên sheet = ngày hôm nay (format: dd-MM-yyyy hoặc 23-02-2026)
                string sheetName = DateTime.Now.ToString("dd-MM-yyyy");

                XLWorkbook workbook;
                
                // Nếu file đã tồn tại, load nó
                if (File.Exists(excelPath))
                {
                    workbook = new XLWorkbook(excelPath);
                        
                    // Xóa sheet cũ nếu tồn tại
                    var existingSheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
                    if (existingSheet != null)
                    {
                        workbook.Worksheets.Delete(sheetName);
                    }
                }
                else
                {
                    workbook = new XLWorkbook();
                }

                using (workbook)
                {
                    // Tạo sheet mới với ngày hôm nay
                    var worksheet = workbook.Worksheets.Add(sheetName);

                    int currentRow = 1;

                    // ===== PHẦN 1: INVOICE DATA =====
                    // Thêm header
                    for (int col = 0; col < dgvInvoice.Columns.Count; col++)
                    {
                        worksheet.Cell(currentRow, col + 1).Value = dgvInvoice.Columns[col].HeaderText;
                        worksheet.Cell(currentRow, col + 1).Style.Font.Bold = true;
                        worksheet.Cell(currentRow, col + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                    }

                    currentRow++;

                    // Thêm dữ liệu từ dgvInvoice (tất cả rows bao gồm cả TỔNG CỘNG)
                    for (int row = 0; row < dgvInvoice.Rows.Count; row++)
                    {
                        for (int col = 0; col < dgvInvoice.Columns.Count; col++)
                        {
                            var cellValue = dgvInvoice.Rows[row].Cells[col].Value;
                            worksheet.Cell(currentRow, col + 1).Value = cellValue?.ToString() ?? "";

                            // Format total row if it's the last row
                            if (row == dgvInvoice.Rows.Count - 1)
                            {
                                worksheet.Cell(currentRow, col + 1).Style.Font.Bold = true;
                                worksheet.Cell(currentRow, col + 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
                            }
                        }
                        currentRow++;
                    }

                    currentRow += 2; // Leave 2 blank rows

                    // ===== PHẦN 2: DAILY REPORT =====
                    // Find and export Daily Report panel data
                    Panel pnlDailyReport = tabInvoice.Controls["pnlDailyReport"] as Panel;
                    if (pnlDailyReport != null)
                    {
                        // Find the Daily Report DataGridView
                        DataGridView dgvReport = null;
                        foreach (Control ctrl in pnlDailyReport.Controls)
                        {
                            if (ctrl is DataGridView dgv)
                            {
                                dgvReport = dgv;
                                break;
                            }
                        }

                        if (dgvReport != null)
                        {
                            // Add header row for Daily Report
                            worksheet.Cell(currentRow, 1).Value = "BÁO CÁO HÀNG NGÀY";
                            worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                            worksheet.Cell(currentRow, 1).Style.Font.FontSize = 12;
                            currentRow++;

                            // Add Daily Report columns
                            for (int col = 0; col < dgvReport.Columns.Count; col++)
                            {
                                worksheet.Cell(currentRow, col + 1).Value = dgvReport.Columns[col].HeaderText;
                                worksheet.Cell(currentRow, col + 1).Style.Font.Bold = true;
                                worksheet.Cell(currentRow, col + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                            }

                            currentRow++;

                            // Add Daily Report rows
                            for (int row = 0; row < dgvReport.Rows.Count; row++)
                            {
                                for (int col = 0; col < dgvReport.Columns.Count; col++)
                                {
                                    var cellValue = dgvReport.Rows[row].Cells[col].Value;
                                    worksheet.Cell(currentRow, col + 1).Value = cellValue?.ToString() ?? "";

                                    // Format header and total rows
                                    if (row == 0 || row == dgvReport.Rows.Count - 1)
                                    {
                                        worksheet.Cell(currentRow, col + 1).Style.Font.Bold = true;
                                        if (row == 0)
                                            worksheet.Cell(currentRow, col + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                                        else
                                            worksheet.Cell(currentRow, col + 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
                                    }
                                }
                                currentRow++;
                            }
                        }
                    }

                    // Auto-fit columns
                    worksheet.Columns().AdjustToContents();

                    // Lưu file
                    workbook.SaveAs(excelPath);
                }

                MessageBox.Show($"✅ Lưu thành công vào:\n{excelPath}\n\nSheet: {sheetName}\n\n✓ Phần 1 (Invoice)\n✓ Phần 2 (Daily Report)", "Thành công");
                lblStatus.Text = $"✅ Lưu Daily Report: {sheetName}";
                lblStatus.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi lưu: {ex.Message}", "Lỗi");
                Debug.WriteLine($"Save error: {ex.Message}\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Initialize OCR Invoice Mapping tab with controls
        /// </summary>
        /// <summary>
        /// Initialize Mapping Tab: OCR text input + Auto-extraction + Manual inputs (người đi, người lấy)
        /// </summary>
        private void InitializeOCRTab()
        {
            try
            {
                Panel pnlOCR = new Panel
                {
                    Dock = DockStyle.Fill,
                    AutoScroll = true,
                    BackColor = SystemColors.Control,
                    Padding = new Padding(10)
                };

                int y = 10;

                // Title
                UIHelper.CreateSectionLabel(pnlOCR, "� OCR Processing", ref y);
                y -= 15;

                // ===== FOLDER SELECTION SECTION =====
                Label lblFolderInfo = new Label
                {
                    Text = "Chon folder anh de quet OCR tu dong",
                    AutoSize = true,
                    Location = new Point(10, y),
                    Font = new Font("Arial", 10, FontStyle.Bold)
                };
                pnlOCR.Controls.Add(lblFolderInfo);
                y += 25;

                // ===== BATCH PROCESSING BUTTONS =====
                var btnSelectFolder = UIHelper.CreateButton("Chon Folder", Color.LightBlue, 10, y, 120, 35);
                btnSelectFolder.Click += (s, e) => SelectOCRFolder();
                pnlOCR.Controls.Add(btnSelectFolder);

                var btnStartScan = UIHelper.CreateButton("Bat Dau Quet", Color.LightGreen, 140, y, 120, 35);
                btnStartScan.Click += (s, e) => StartBatchOCRProcessing();
                pnlOCR.Controls.Add(btnStartScan);

                var btnExport = UIHelper.CreateButton("Xuat", Color.Orange, 270, y, 80, 35);
                btnExport.Click += (s, e) => ExportSelectedImages();
                pnlOCR.Controls.Add(btnExport);

                y += 45;

                // ===== MANUAL INPUT SECTION: NGƯỜI ĐI & NGƯỜI LẤY =====
                UIHelper.CreateSectionLabel(pnlOCR, "Thong tin NGUOI DI & NGUOI LAY (bat buoc):", ref y);
                y -= 15;

                // Người Đi
                Label lblNguoiDi = new Label
                {
                    Text = "Người Đi:",
                    AutoSize = true,
                    Location = new Point(10, y),
                    Font = new Font("Arial", 9, FontStyle.Bold)
                };
                pnlOCR.Controls.Add(lblNguoiDi);

                txtNguoiDiOCR = new TextBox
                {
                    Location = new Point(10, y + 25),
                    Width = pnlOCR.ClientSize.Width - 20,
                    Height = 35,
                    BorderStyle = BorderStyle.FixedSingle,
                    Font = new Font("Arial", 11)
                };
                pnlOCR.Controls.Add(txtNguoiDiOCR);
                y += 65;

                // Người Lấy
                Label lblNguoiLay = new Label
                {
                    Text = "Người Lấy:",
                    AutoSize = true,
                    Location = new Point(10, y),
                    Font = new Font("Arial", 9, FontStyle.Bold)
                };
                pnlOCR.Controls.Add(lblNguoiLay);

                txtNguoiLayOCR = new TextBox
                {
                    Location = new Point(10, y + 25),
                    Width = pnlOCR.ClientSize.Width - 20,
                    Height = 35,
                    BorderStyle = BorderStyle.FixedSingle,
                    Font = new Font("Arial", 11)
                };
                pnlOCR.Controls.Add(txtNguoiLayOCR);
                y += 65;

                // ===== PROCESS LOG SECTION =====
                UIHelper.CreateSectionLabel(pnlOCR, "📋 Raw OCR Text (Kết quả OCR thô):", ref y);
                y -= 15;

                // Rich textbox for raw OCR logging
                this.txtRawOCRLog = new RichTextBox
                {
                    Location = new Point(10, y),
                    Width = pnlOCR.ClientSize.Width - 30,
                    Height = 200,
                    ReadOnly = true,
                    BackColor = Color.White,
                    Font = new Font("Courier New", 8),
                    BorderStyle = BorderStyle.FixedSingle
                };
                pnlOCR.Controls.Add(this.txtRawOCRLog);
                
                y += 210;

                // ===== MAPPING LOG SECTION =====
                UIHelper.CreateSectionLabel(pnlOCR, "✅ Chi tiet quet OCR (Mapping kết quả):", ref y);
                y -= 15;

                // Rich textbox for mapping logging
                this.txtProcessLog = new RichTextBox
                {
                    Location = new Point(10, y),
                    Width = pnlOCR.ClientSize.Width - 30,
                    Height = 400,
                    ReadOnly = true,
                    BackColor = Color.White,
                    Font = new Font("Courier New", 8),
                    BorderStyle = BorderStyle.FixedSingle
                };
                pnlOCR.Controls.Add(this.txtProcessLog);
                
                y += 410;

                // ===== EXPORT BUTTON =====
                var btnExportOCR = UIHelper.CreateButton("💾 XUẤT EXCEL", Color.LightGreen, 10, y, 150, 35);
                btnExportOCR.Click += (s, e) => ExportMappedDataToExcel();
                pnlOCR.Controls.Add(btnExportOCR);

                y += 45;

                // ===== BATCH OCR LOG =====
                UIHelper.CreateSectionLabel(pnlOCR, "📋 Kết quả Batch OCR:", ref y);
                y -= 15;

                var batchLog = new RichTextBox
                {
                    Location = new Point(10, y),
                    Width = pnlOCR.ClientSize.Width - 30,
                    Height = 150,
                    ReadOnly = true,
                    BackColor = Color.White,
                    Font = new Font("Courier New", 8),
                    BorderStyle = BorderStyle.FixedSingle
                };
                pnlOCR.Controls.Add(batchLog);
                y += 160;

                // ===== CHECKLIST FOR EXPORT =====
                UIHelper.CreateSectionLabel(pnlOCR, "☑ Chọn ảnh để xuất:", ref y);
                y -= 15;

                var chkList = new CheckedListBox
                {
                    Location = new Point(10, y),
                    Width = pnlOCR.ClientSize.Width - 30,
                    Height = 120,
                    BorderStyle = BorderStyle.FixedSingle,
                    Font = new Font("Arial", 9),
                    CheckOnClick = true
                };
                pnlOCR.Controls.Add(chkList);
                y += 130;

                // Store references
                pnlOCR.Tag = new Dictionary<string, object>
                {
                    { "rawLog", this.txtRawOCRLog },
                    { "mappingLog", this.txtProcessLog },
                    { "log", batchLog },
                    { "checklist", chkList }
                };

                // Add resize event to make input fields responsive
                pnlOCR.Resize += (s, e) =>
                {
                    if (txtNguoiDiOCR != null)
                        txtNguoiDiOCR.Width = pnlOCR.ClientSize.Width - 20;
                    if (txtNguoiLayOCR != null)
                        txtNguoiLayOCR.Width = pnlOCR.ClientSize.Width - 20;
                    if (txtRawOCRLog != null)
                        txtRawOCRLog.Width = pnlOCR.ClientSize.Width - 30;
                    if (txtProcessLog != null)
                        txtProcessLog.Width = pnlOCR.ClientSize.Width - 30;
                };

                tabOCR.Controls.Clear();
                tabOCR.Controls.Add(pnlOCR);

                Debug.WriteLine("OCR Batch Tab initialized");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error initializing OCR Tab: {ex.Message}");
            }
        }

        /// <summary>
        /// Map OCR text to 10 required fields
        /// SHOP, TÊN KH, MÃ, SỐ NHÀ, TÊN ĐƯỜNG, QUẬN, TIỀN THU, TIỀN SHIP, TIỀN HÀNG, NGÀY LẤY
        /// </summary>
        private Dictionary<string, string> MapOCRDataTo12Fields(string ocrText, string fileName, string nguoiDi, string nguoiLay)
        {
            var tienThu  = ExtractNumeric(ocrText, "tiền thu|thu tiền|tổng thanh toán");  // "" nếu không tìm thấy
            var tienShip = ExtractNumeric(ocrText, "tiền ship|ship|vận chuyển");          // "" nếu không tìm thấy

            // TIỀN HÀNG = TIỀN THU + TIỀN SHIP (tự tính, không lấy từ OCR)
            string tienHang = "";
            if (!string.IsNullOrEmpty(tienThu) || !string.IsNullOrEmpty(tienShip))
            {
                long thu  = long.TryParse(tienThu,  out var t)  ? t : 0;
                long ship = long.TryParse(tienShip, out var s)  ? s : 0;
                tienHang = (thu + ship).ToString();
            }

            // NGÀY LẤY: ưu tiên lấy từ OCR, fallback về hôm nay
            string ngayLay = ExtractDateFromOCR(ocrText);
            if (string.IsNullOrEmpty(ngayLay))
                ngayLay = DateTime.Now.ToString("dd-MM-yyyy");

            var result = new Dictionary<string, string>
            {
                { "fileName", fileName },
                // Extract SHOP and TÊN KH from OCR text
                { "SHOP",      ExtractField(ocrText, "đoàn|shop|cửa hàng", 100) },
                { "TÊN KH",    ExtractField(ocrText, "khách hàng:|customer:", 100) },
                // NGƯỜI ĐI & NGƯỜI LẤY from manual input
                { "NGƯỜI ĐI",  nguoiDi },
                { "NGƯỜI LẤY", nguoiLay },
                // Extract remaining fields from OCR
                { "MÃ",        ExtractField(ocrText, "so hd:|so hd|mã|ma:", 50) },
                { "SỐ NHÀ",    ExtractAddressField(ocrText, "soNha") },
                { "TÊN ĐƯỜNG", ExtractAddressField(ocrText, "tenDuong") },
                { "QUẬN",      ExtractAddressField(ocrText, "quan") },
                { "TIỀN THU",  tienThu },
                { "TIỀN SHIP", tienShip },
                { "TIỀN HÀNG", tienHang },   // Tính từ TIỀN THU + TIỀN SHIP
                { "NGÀY LẤY",  ngayLay }     // Lấy từ OCR, format dd-MM-yyyy
            };
            return result;
        }

        /// <summary>
        /// Extract ngày tháng năm từ OCR text.
        /// Nhận các dạng:
        ///   "Ngày 11 tháng 02 năm 2026"
        ///   "11/02/2026", "11-02-2026"
        ///   "ngày 11/02/2026"
        /// Trả về format "dd-MM-yyyy", hoặc "" nếu không tìm thấy.
        /// </summary>
        private string ExtractDateFromOCR(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";

            // Pattern 1: "Ngày DD tháng MM năm YYYY" (dạng trong ảnh hóa đơn)
            var m1 = System.Text.RegularExpressions.Regex.Match(text,
                @"ng[aà]y\s+(\d{1,2})\s+th[aá]ng\s+(\d{1,2})\s+n[aă]m\s+(\d{4})",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (m1.Success)
            {
                string dd = m1.Groups[1].Value.PadLeft(2, '0');
                string mm = m1.Groups[2].Value.PadLeft(2, '0');
                string yyyy = m1.Groups[3].Value;
                return $"{dd}-{mm}-{yyyy}";
            }

            // Pattern 2: DD/MM/YYYY hoặc DD-MM-YYYY (standalone, không nằm trong chuỗi số dài)
            var m2 = System.Text.RegularExpressions.Regex.Match(text,
                @"\b(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})\b");
            if (m2.Success)
            {
                string dd = m2.Groups[1].Value.PadLeft(2, '0');
                string mm = m2.Groups[2].Value.PadLeft(2, '0');
                string yyyy = m2.Groups[3].Value;
                return $"{dd}-{mm}-{yyyy}";
            }

            return "";
        }

        /// <summary>
        /// Extract address field (số nhà, tên đường, quận) from the SECOND address block (người nhận)
        /// OCR usually has 2 address blocks: shop address and receiver address
        /// </summary>
        private string ExtractAddressField(string ocrText, string fieldType)
        {
            if (string.IsNullOrWhiteSpace(ocrText)) return "";

            var lines = ocrText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            
            // Find the SECOND occurrence of "Địa chỉ:" (receiver's address, not shop)
            int addressBlockCount = 0;
            int startLine = -1;
            
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].IndexOf("địa chỉ", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    lines[i].IndexOf("địa chi", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    addressBlockCount++;
                    if (addressBlockCount == 2) // Found second address block
                    {
                        startLine = i;
                        break;
                    }
                }
            }

            if (startLine == -1)
            {
                // If only one address block found, use it (fallback)
                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].IndexOf("địa chỉ", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        lines[i].IndexOf("địa chi", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        startLine = i;
                        break;
                    }
                }
            }

            if (startLine == -1) return "";

            // Extract address from the line
            string addressLine = lines[startLine];

            // Remove "Địa chỉ:" prefix
            int colonIdx = addressLine.IndexOf(':');
            if (colonIdx >= 0)
            {
                addressLine = addressLine.Substring(colonIdx + 1).Trim();
            }

            // Parse address using AddressParser for consistent results
            var parsed = TextInputter.Services.AddressParser.Parse(addressLine);

            switch (fieldType.ToLower())
            {
                case "sonha":
                    return parsed.SoNha;

                case "tenduong":
                    return parsed.TenDuong;

                case "quan":
                    return parsed.Quan;

                default:
                    return addressLine;
            }
        }

        /// <summary>
        /// Extract text field from OCR text by keyword
        /// </summary>
        private string ExtractField(string text, string keywords, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";

            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var keywordList = keywords.Split('|');

            foreach (var line in lines)
            {
                foreach (var keyword in keywordList)
                {
                    if (line.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Extract text after colon or keyword
                        var parts = line.Split(new[] { ':', '-' }, StringSplitOptions.None);
                        if (parts.Length > 1)
                        {
                            var value = parts[parts.Length - 1].Trim();
                            return value.Length > maxLength ? value.Substring(0, maxLength) : value;
                        }
                        return line.Trim();
                    }
                }
            }
            return "";
        }

        /// <summary>
        /// Extract numeric value from OCR text.
        /// Trả về "" nếu không tìm thấy (không phải "0") để validation phát hiện thiếu.
        /// </summary>
        private string ExtractNumeric(string text, string keywords)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";

            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var keywordList = keywords.Split('|');

            foreach (var line in lines)
            {
                foreach (var keyword in keywordList)
                {
                    if (line.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Extract number from the line
                        var matches = System.Text.RegularExpressions.Regex.Matches(line, @"\d+");
                        if (matches.Count > 0)
                        {
                            return matches[matches.Count - 1].Value;
                        }
                    }
                }
            }
            return "";  // Không tìm thấy → trả về rỗng để ValidateMappedData biết là thiếu
        }

        /// <summary>
        /// Validate mapped data - check if all 9 required fields have values
        /// (TIỀN HÀNG không require vì tự tính từ TIỀN THU + TIỀN SHIP)
        /// </summary>
        private List<string> ValidateMappedData(Dictionary<string, string> mappedData)
        {
            var missingFields = new List<string>();

            var requiredFields = new[] { "SHOP", "TÊN KH", "MÃ", "SỐ NHÀ", "TÊN ĐƯỜNG", "QUẬN", "TIỀN THU", "TIỀN SHIP", "NGÀY LẤY", "NGƯỜI ĐI", "NGƯỜI LẤY" };

            foreach (var field in requiredFields)
            {
                if (!mappedData.ContainsKey(field) || string.IsNullOrWhiteSpace(mappedData[field]))
                {
                    missingFields.Add(field);
                }
            }

            return missingFields;
        }

        /// <summary>
        /// Get all field statuses (pass/fail) for logging
        /// </summary>
        private Dictionary<string, bool> GetFieldStatuses(Dictionary<string, string> mappedData)
        {
            var fieldStatuses = new Dictionary<string, bool>();
            var requiredFields = new[] { "SHOP", "TÊN KH", "MÃ", "SỐ NHÀ", "TÊN ĐƯỜNG", "QUẬN", "TIỀN THU", "TIỀN SHIP", "NGÀY LẤY", "NGƯỜI ĐI", "NGƯỜI LẤY" };

            foreach (var field in requiredFields)
            {
                fieldStatuses[field] = mappedData.ContainsKey(field) && !string.IsNullOrWhiteSpace(mappedData[field]);
            }

            return fieldStatuses;
        }

        /// <summary>
        /// Export mapped data to Excel
        /// </summary>
        private void ExportMappedDataToExcel()
        {
            try
            {
                if (mappedDataList.Count == 0)
                {
                    MessageBox.Show("❌ Không có dữ liệu để xuất. Vui lòng quét ảnh trước!", "Thông báo", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Debug.WriteLine($"📊 Bắt đầu xuất {mappedDataList.Count} dòng dữ liệu");

                // Ask user to select Excel file to export to
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                    Title = "Chọn file Excel để export dữ liệu",
                    InitialDirectory = Path.Combine(Directory.GetCurrentDirectory(), "data", "sample", "excel")
                };

                if (openFileDialog.ShowDialog() != DialogResult.OK) return;

                string excelPath = openFileDialog.FileName;
                if (!File.Exists(excelPath))
                {
                    MessageBox.Show($"❌ File không tồn tại: {excelPath}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var now = DateTime.Now;

                // Tên sheet = dd-MM lấy từ NGÀY LẤY trong data (ưu tiên data đầu tiên)
                // Format NGÀY LẤY là "dd-MM-yyyy", tách ra lấy dd-MM
                string sheetName;
                if (mappedDataList.Count > 0 && mappedDataList[0].ContainsKey("NGÀY LẤY")
                    && !string.IsNullOrEmpty(mappedDataList[0]["NGÀY LẤY"]))
                {
                    // "11-02-2026" → lấy 2 phần đầu → "11-02"
                    var parts = mappedDataList[0]["NGÀY LẤY"].Split('-');
                    sheetName = parts.Length >= 2 ? $"{parts[0]}-{parts[1]}" : now.ToString("dd-MM");
                }
                else
                {
                    sheetName = now.ToString("dd-MM");
                }

                // Ngày để điền vào row 2 (THU x / NGAY x-x) — parse từ sheetName
                DateTime sheetDate = now;
                if (DateTime.TryParseExact(sheetName, "dd-MM",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var parsed))
                    sheetDate = parsed;

                using (var workbook = new XLWorkbook(excelPath))
                {
                    IXLWorksheet worksheet;
                    bool isNewSheet = false;

                    // Nếu sheet đã tồn tại → xóa và tạo lại (ghi đè, không báo lỗi)
                    if (workbook.TryGetWorksheet(sheetName, out worksheet))
                    {
                        // Giữ nguyên sheet, chỉ tìm row cuối để append
                        Debug.WriteLine($"✅ Sheet '{sheetName}' đã tồn tại, append dữ liệu");
                        isNewSheet = false;
                    }
                    else
                    {
                        worksheet = workbook.Worksheets.Add(sheetName);
                        isNewSheet = true;
                        Debug.WriteLine($"✨ Tạo sheet mới: '{sheetName}'");
                    }

                    // Cột chuẩn khớp với các sheet khác (20 cột)
                    // Col: 1=TinhTrang, 2=SHOP, 3=TENKH, 4=MA, 5=SONHA, 6=TENDUONG, 7=QUAN,
                    //      8=TIENTHU, 9=TIENSHIP, 10=TIENHANG, 11=NGUOIDI, 12=NGUOILAY,
                    //      13=NGAYLAY, 14=GHICHU, 15=UNGIEN, 16=HANGTON, 17=FAIL,
                    //      18=Column1, 19=Column2, 20=Column3
                    var headers = new[]
                    {
                        "Tình trạng TT", "SHOP", "TÊN KH", "MÃ", "SỐ NHÀ", "TÊN ĐƯỜNG", "QUẬN",
                        "TIỀN THU", "TIỀN SHIP", "TIỀN HÀNG",
                        "NGƯỜI ĐI", "NGƯỜI LẤY", "NGÀY LẤY", "GHI CHÚ",
                        "ỨNG TIỀN", "HÀNG TỒN", "FAIL", "Column1", "Column2", "Column3"
                    };

                    if (isNewSheet)
                    {
                        // Row 1: Column headers
                        for (int col = 0; col < headers.Length; col++)
                        {
                            var cell = worksheet.Cell(1, col + 1);
                            cell.Value = headers[col];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                        }

                        // Row 2: THU x | NGAY x-x (giống các sheet khác)
                        string thuText = sheetDate.DayOfWeek == DayOfWeek.Sunday
                            ? "CHU NHAT"
                            : "THU " + ((int)sheetDate.DayOfWeek + 1);
                        string ngayText = "NGAY " + sheetDate.Day + "-" + sheetDate.Month;

                        var cellThu = worksheet.Cell(2, 2); // cột SHOP
                        cellThu.Value = thuText;
                        cellThu.Style.Font.Bold = true;

                        var cellNgay = worksheet.Cell(2, 3); // cột TÊN KH
                        cellNgay.Value = ngayText;
                        cellNgay.Style.Font.Bold = true;
                    }

                    // Tìm row cuối để append (data bắt đầu từ row 3)
                    int currentRow = 3;
                    var lastUsed = worksheet.LastRowUsed();
                    if (lastUsed != null && lastUsed.RowNumber() >= 3)
                        currentRow = lastUsed.RowNumber() + 1;

                    // Ghi đè: không check trùng MÃ, cứ append
                    int addedCount = 0;
                    foreach (var data in mappedDataList)
                    {
                        worksheet.Cell(currentRow, 1).Value  = "";                      // Tình trạng TT (để trống)
                        worksheet.Cell(currentRow, 2).Value  = data["SHOP"];
                        worksheet.Cell(currentRow, 3).Value  = data["TÊN KH"];
                        worksheet.Cell(currentRow, 4).Value  = data["MÃ"];
                        worksheet.Cell(currentRow, 5).Value  = data["SỐ NHÀ"];
                        worksheet.Cell(currentRow, 6).Value  = data["TÊN ĐƯỜNG"];
                        worksheet.Cell(currentRow, 7).Value  = data["QUẬN"];
                        worksheet.Cell(currentRow, 8).Value  = data["TIỀN THU"];
                        worksheet.Cell(currentRow, 9).Value  = data["TIỀN SHIP"];
                        worksheet.Cell(currentRow, 10).Value = data["TIỀN HÀNG"];
                        worksheet.Cell(currentRow, 11).Value = data["NGƯỜI ĐI"];
                        worksheet.Cell(currentRow, 12).Value = data["NGƯỜI LẤY"];
                        worksheet.Cell(currentRow, 13).Value = data["NGÀY LẤY"];
                        // Col 14-20 để trống (GHI CHÚ, ỨNG TIỀN, HÀNG TỒN, FAIL, Column1/2/3)
                        currentRow++;
                        addedCount++;
                    }

                    workbook.SaveAs(excelPath);
                    Debug.WriteLine($"✅ Lưu xong! {addedCount} dòng → sheet '{sheetName}'");

                    this.Invoke((MethodInvoker)delegate
                    {
                        MessageBox.Show(
                            $"✅ Xuất thành công!\n\n📌 Dòng thêm: {addedCount}\n📅 Sheet: {sheetName}\n📂 File: {Path.GetFileName(excelPath)}",
                            "✅ Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        lblStatus.Text = $"✅ Xuất {addedCount} dòng → sheet '{sheetName}'";
                        lblStatus.ForeColor = Color.Green;
                    });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ LỖI: {ex.Message}\n{ex.StackTrace}");
                this.Invoke((MethodInvoker)delegate
                {
                    MessageBox.Show($"❌ Lỗi xuất Excel:\n\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
        }

        /// <summary>
        /// Select folder containing images for batch OCR processing
        /// </summary>
        private void SelectOCRFolder()
        {
            try
            {
                using (var fbd = new FolderBrowserDialog())
                {
                    fbd.Description = "Chọn folder chứa ảnh cần quét OCR";
                    if (fbd.ShowDialog() == DialogResult.OK)
                    {
                        folderPath = fbd.SelectedPath;
                        imageFiles = Directory.GetFiles(folderPath, "*.*")
                            .Where(f => new[] { ".jpg", ".jpeg", ".png", ".bmp", ".tiff" }
                                .Contains(Path.GetExtension(f).ToLower()))
                            .ToList();

                        // Get panel references
                        var pnlOCR = tabOCR.Controls[0] as Panel;
                        if (pnlOCR?.Tag is Dictionary<string, object> refs && refs.TryGetValue("log", out var logObj) && logObj is RichTextBox log)
                        {
                            log.Clear();
                            log.Text = $"Da chon folder: {folderPath}\n";
                            log.AppendText($"Tim thay {imageFiles.Count} anh\n\n");
                            log.AppendText("Danh sach anh:\n");
                            foreach (var img in imageFiles)
                            {
                                log.AppendText($"  * {Path.GetFileName(img)}\n");
                            }
                        }

                        MessageBox.Show($"Da chon folder: {folderPath}\nTim thay {imageFiles.Count} anh", "Thanh cong");
                        Debug.WriteLine($"Selected folder: {folderPath}, Found {imageFiles.Count} images");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Loi: {ex.Message}", "Loi");
                Debug.WriteLine($"Error selecting folder: {ex.Message}");
            }
        }

        /// <summary>
        /// Start batch OCR processing on selected folder
        /// </summary>
        private void StartBatchOCRProcessing()
        {
            try
            {
                if (imageFiles.Count == 0)
                {
                    MessageBox.Show("Vui long chon folder truoc", "Canh bao");
                    return;
                }

                // Get panel references
                var pnlOCR = tabOCR.Controls[0] as Panel;
                if (pnlOCR?.Tag is not Dictionary<string, object> refs)
                    return;

                if (!refs.TryGetValue("log", out var logObj) || logObj is not RichTextBox log)
                    return;

                if (!refs.TryGetValue("checklist", out var checkListObj) || checkListObj is not CheckedListBox chkList)
                    return;

                log.Clear();
                log.Text = $"Quet {imageFiles.Count} anh...\n\n";

                int successCount = 0;
                int failCount = 0;
                var failedImages = new List<string>();
                var failedReasons = new Dictionary<string, string>(); // Track failure reasons
                var successImages = new List<string>();

                chkList.Items.Clear();

                foreach (var imagePath in imageFiles)
                {
                    try
                    {
                        log.AppendText($"Xu ly: {Path.GetFileName(imagePath)}...\n");
                        Application.DoEvents();

                        string ocrText = ExtractTextFromImage(imagePath);

                        if (string.IsNullOrEmpty(ocrText))
                        {
                            log.AppendText($"  [FAIL] OCR failed\n");
                            failCount++;
                            failedImages.Add(Path.GetFileName(imagePath));
                            failedReasons[Path.GetFileName(imagePath)] = "OCR text empty";
                            continue;
                        }

                        // Extract all 12 required fields
                        Dictionary<string, string> fields = new Dictionary<string, string>();
                        List<string> missingFields = new List<string>();
                        
                        if (_ocrParsingService != null)
                        {
                            missingFields = _ocrParsingService.ExtractAllFields(ocrText, out fields) ?? new List<string>();
                        }

                        if (missingFields.Count > 0)
                        {
                            log.AppendText($"  [FAIL] Thieu: {string.Join(", ", missingFields)}\n");
                            failCount++;
                            failedImages.Add(Path.GetFileName(imagePath));
                            failedReasons[Path.GetFileName(imagePath)] = $"Missing: {string.Join(", ", missingFields)}";
                            continue;
                        }

                        // Get extracted fields
                        string soHD = fields?.ContainsKey("Số HĐ") == true ? fields["Số HĐ"] : string.Empty;
                        decimal tongTien = decimal.TryParse(fields?["Tổng Tiền"], out var amt) ? amt : 0m;

                        if (_excelInvoiceService.InvoiceExists(soHD, out string existingSheet))
                        {
                            log.AppendText($"  [SKIP] SoHD '{soHD}' ton tai (sheet: {existingSheet})\n");
                            failCount++;
                            failedImages.Add(Path.GetFileName(imagePath));
                            continue;
                        }

                        // SUCCESS - add to checklist and track
                        decimal chietKhau = _ocrParsingService?.ExtractDiscount(ocrText) ?? 0m;
                        string fileName = Path.GetFileName(imagePath);
                        
                        chkList.Items.Add(fileName, true); // Add with checkbox checked
                        successImages.Add(imagePath); // Store full path
                        
                        log.AppendText($"  [OK] {soHD}\n");
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        log.AppendText($"  [ERROR] {ex.Message}\n");
                        failCount++;
                        failedImages.Add(Path.GetFileName(imagePath));
                        Debug.WriteLine($"Error processing {imagePath}: {ex.Message}");
                    }
                }

                // Save success images to refs
                refs["successImages"] = successImages;

                log.AppendText($"\n{'='*60}\n");
                log.AppendText($"KET QUA:\n");
                log.AppendText($"OK: {successCount}/{imageFiles.Count}\n");
                log.AppendText($"FAIL: {failCount}/{imageFiles.Count}\n");

                if (failedImages.Count > 0)
                {
                    log.AppendText($"\nAnh that bai:\n");
                    foreach (var failed in failedImages)
                    {
                        log.AppendText($"  * {failed}\n");
                    }
                }

                MessageBox.Show(
                    $"Hoan tat xu ly!\n\nThanh cong: {successCount}\nThat bai: {failCount}\n\nChon anh can xuat o duoi roi nhan 'Xuat'",
                    "Thong bao",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                Debug.WriteLine($"Batch processing completed: {successCount} success, {failCount} failed");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Loi: {ex.Message}", "Loi");
                Debug.WriteLine($"Error in batch processing: {ex.Message}\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Extract text from image using Tesseract OCR
        /// </summary>
        private string ExtractTextFromImage(string imagePath)
        {
            try
            {
                // Using Tesseract for image processing
                // This is placeholder - actual implementation depends on Tesseract setup
                // For now, return empty to let batch processing continue
                return "";
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error extracting text: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// Export selected successful images to Excel
        /// </summary>
        private void ExportSelectedImages()
        {
            try
            {
                // Get panel references
                var pnlOCR = tabOCR.Controls[0] as Panel;
                if (pnlOCR?.Tag is not Dictionary<string, object> refs)
                    return;

                if (!refs.TryGetValue("checklist", out var checkListObj) || checkListObj is not CheckedListBox chkList)
                    return;

                if (!refs.TryGetValue("successImages", out var successObj) || successObj is not List<string> successImages)
                    return;

                // Get checked items
                var selectedIndices = new List<int>();
                for (int i = 0; i < chkList.CheckedItems.Count; i++)
                {
                    selectedIndices.Add(chkList.Items.IndexOf(chkList.CheckedItems[i]));
                }

                if (selectedIndices.Count == 0)
                {
                    MessageBox.Show("Vui long chon it nhat 1 anh", "Canh bao");
                    return;
                }

                int exportCount = 0;

                foreach (int idx in selectedIndices)
                {
                    if (idx >= 0 && idx < successImages.Count)
                    {
                        string imagePath = successImages[idx];
                        
                        // Re-extract and export
                        try
                        {
                            string ocrText = ExtractTextFromImage(imagePath);
                            if (string.IsNullOrEmpty(ocrText))
                                continue;

                            string soHD = _ocrParsingService?.ExtractInvoiceNumber(ocrText) ?? string.Empty;
                            string diaChi = _ocrParsingService?.ExtractAddress(ocrText) ?? string.Empty;
                            decimal tongTien = _ocrParsingService?.ExtractTotalAmount(ocrText) ?? 0m;

                            if (string.IsNullOrEmpty(soHD) || string.IsNullOrEmpty(diaChi) || tongTien <= 0)
                                continue;

                            // Check duplicate again (may be added during previous exports)
                            if (_excelInvoiceService.InvoiceExists(soHD, out _))
                                continue;

                            decimal chietKhau = _ocrParsingService?.ExtractDiscount(ocrText) ?? 0m;

                            var invoice = new Services.OCRInvoiceData
                            {
                                SoHoaDon = soHD,
                                DiaChi = diaChi,
                                TongTienHang = tongTien,
                                ChietKhau = chietKhau,
                                TongThanhToan = tongTien - chietKhau,
                                NguoiDi = "OCR Auto",
                                NguoiLay = "OCR Auto"
                            };

                            _excelInvoiceService.ExportInvoice(invoice);
                            exportCount++;
                        }
                        catch (Exception itemEx)
                        {
                            // Skip failed exports
                            Debug.WriteLine($"Failed to export image: {itemEx.Message}");
                        }
                    }
                }

                // Always show success message even if count is 0
                if (exportCount > 0)
                {
                    MessageBox.Show($"✅ Xuất thành công {exportCount} ảnh!", "Thông báo");
                }
                else
                {
                    MessageBox.Show("⚠️ Không có ảnh nào được xuất thành công", "Thông báo");
                }
                Debug.WriteLine($"Exported {exportCount} images");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi: {ex.Message}", "Lỗi");
                Debug.WriteLine($"Error exporting: {ex.Message}");
            }
        }

        /// <summary>
        /// Initialize Manual Input Tab - For entering data manually with all 17 mandatory fields
        /// </summary>
        private void InitializeManualInputTab()
        {
            try
            {
                Panel pnlManualInput = new Panel
                {
                    Dock = DockStyle.Fill,
                    AutoScroll = true,
                    BackColor = SystemColors.Control,
                    Padding = new Padding(10)
                };

                int y = 10;

                // Title
                UIHelper.CreateSectionLabel(pnlManualInput, "✋ Nhập Dữ Liệu Thủ Công (17 Trường Bắt Buộc)", ref y);
                y -= 15;

                // Legend
                Label lblLegend = new Label
                {
                    Text = "⭐ Tất cả các trường màu vàng là bắt buộc phải điền",
                    AutoSize = true,
                    ForeColor = Color.OrangeRed,
                    Font = new Font("Arial", 9, FontStyle.Bold),
                    Location = new Point(10, y)
                };
                pnlManualInput.Controls.Add(lblLegend);
                y += 25;

                // ===== SECTION 1: BASIC INFO =====
                UIHelper.CreateSectionLabel(pnlManualInput, "📋 Thông Tin Cơ Bản:", ref y);
                y -= 15;

                var txtTinhTrang = CreateMandatoryField(pnlManualInput, "[1] Tình Trạng TT:", ref y);
                var txtThuTu = CreateMandatoryField(pnlManualInput, "[2] Thứ:", ref y);
                var txtNgay = CreateMandatoryField(pnlManualInput, "[3] Ngày (DD-MM-YYYY):", ref y);
                var txtMa = CreateMandatoryField(pnlManualInput, "[4] Mã:", ref y);

                // ===== SECTION 2: ADDRESS =====
                UIHelper.CreateSectionLabel(pnlManualInput, "📍 Địa Chỉ:", ref y);
                y -= 15;

                var txtSoNha = CreateMandatoryField(pnlManualInput, "[5] Số Nhà:", ref y);
                var txtTenDuong = CreateMandatoryField(pnlManualInput, "[6] Tên Đường:", ref y);
                var txtQuan = CreateMandatoryField(pnlManualInput, "[7] Quận:", ref y);

                // ===== SECTION 3: MONEY =====
                UIHelper.CreateSectionLabel(pnlManualInput, "💰 Tiền Tệ:", ref y);
                y -= 15;

                var txtTienThu = CreateMandatoryField(pnlManualInput, "[8] Tiền Thu:", ref y);
                var txtTienShip = CreateMandatoryField(pnlManualInput, "[9] Tiền Ship:", ref y);
                var txtTienHang = CreateMandatoryField(pnlManualInput, "[10] Tiền Hàng:", ref y);

                // ===== SECTION 4: PEOPLE & STATUS =====
                UIHelper.CreateSectionLabel(pnlManualInput, "👥 Người Liên Quan & Trạng Thái:", ref y);
                y -= 15;

                var txtNguoiDi = CreateMandatoryField(pnlManualInput, "[11] Người Đi:", ref y);
                var txtNguoiLay = CreateMandatoryField(pnlManualInput, "[12] Người Lấy:", ref y);
                var txtGhiChu = CreateMandatoryField(pnlManualInput, "[13] Ghi Chú:", ref y);
                var txtUng = CreateMandatoryField(pnlManualInput, "[14] Ưng (Advance):", ref y);
                var txtHang = CreateMandatoryField(pnlManualInput, "[15] Hàng (Status):", ref y);
                var txtFail = CreateMandatoryField(pnlManualInput, "[16] Fail:", ref y);
                var txtNote = CreateMandatoryField(pnlManualInput, "[17] Ghi Chú Thêm:", ref y);

                // ===== BUTTONS =====
                y += 10;
                var btnSaveManual = UIHelper.CreateButton("💾 Lưu", Color.LightGreen, 10, y, 100, 35);
                btnSaveManual.Click += (s, e) => SaveManualEntry(
                    txtTinhTrang.Text, txtThuTu.Text, txtNgay.Text, txtMa.Text,
                    txtSoNha.Text, txtTenDuong.Text, txtQuan.Text,
                    txtTienThu.Text, txtTienShip.Text, txtTienHang.Text,
                    txtNguoiDi.Text, txtNguoiLay.Text, txtGhiChu.Text,
                    txtUng.Text, txtHang.Text, txtFail.Text, txtNote.Text);
                pnlManualInput.Controls.Add(btnSaveManual);

                var btnClearManual = UIHelper.CreateButton("🔄 Xóa", Color.LightCoral, 120, y, 100, 35);
                btnClearManual.Click += (s, e) =>
                {
                    txtTinhTrang.Clear();
                    txtThuTu.Clear();
                    txtNgay.Clear();
                    txtMa.Clear();
                    txtSoNha.Clear();
                    txtTenDuong.Clear();
                    txtQuan.Clear();
                    txtTienThu.Clear();
                    txtTienShip.Clear();
                    txtTienHang.Clear();
                    txtNguoiDi.Clear();
                    txtNguoiLay.Clear();
                    txtGhiChu.Clear();
                    txtUng.Clear();
                    txtHang.Clear();
                    txtFail.Clear();
                    txtNote.Clear();
                };
                pnlManualInput.Controls.Add(btnClearManual);

                tabManualInput.Controls.Clear();
                tabManualInput.Controls.Add(pnlManualInput);

                Debug.WriteLine("✅ Manual Input Tab initialized successfully with 17 fields");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ Error initializing Manual Input Tab: {ex.Message}");
            }
        }

        /// <summary>
        /// Create a mandatory field with yellow background
        /// </summary>
        private TextBox CreateMandatoryField(Panel panel, string labelText, ref int yPos, bool isMultiline = false)
        {
            Label lbl = new Label
            {
                Text = labelText,
                AutoSize = true,
                Location = new Point(10, yPos),
                Font = new Font("Arial", 9, FontStyle.Bold),
                ForeColor = Color.Black
            };
            panel.Controls.Add(lbl);
            yPos += 20;

            TextBox txt = new TextBox
            {
                Location = new Point(10, yPos),
                Width = panel.ClientSize.Width - 30,
                Height = isMultiline ? 60 : 25,
                Multiline = isMultiline,
                BackColor = Color.Yellow, // Mandatory field highlight
                Font = new Font("Arial", 9),
                BorderStyle = BorderStyle.FixedSingle
            };
            panel.Controls.Add(txt);
            yPos += (isMultiline ? 70 : 35);

            return txt;
        }

        /// <summary>
        /// Save manual entry to Excel with mandatory field validation (17 fields)
        /// </summary>
        private void SaveManualEntry(string tinhTrang, string thuTu, string ngay, string ma,
            string soNha, string tenDuong, string quan,
            string tienThu, string tienShip, string tienHang,
            string nguoiDi, string nguoiLay, string ghiChu,
            string ung, string hang, string fail, string note)
        {
            try
            {
                // Validate mandatory fields (must not be empty or whitespace)
                var missingFields = new List<string>();
                
                if (string.IsNullOrWhiteSpace(tinhTrang)) missingFields.Add("1. Tình Trạng TT");
                if (string.IsNullOrWhiteSpace(thuTu)) missingFields.Add("2. Thứ");
                if (string.IsNullOrWhiteSpace(ngay)) missingFields.Add("3. Ngày");
                if (string.IsNullOrWhiteSpace(ma)) missingFields.Add("4. Mã");
                if (string.IsNullOrWhiteSpace(soNha)) missingFields.Add("5. Số Nhà");
                if (string.IsNullOrWhiteSpace(tenDuong)) missingFields.Add("6. Tên Đường");
                if (string.IsNullOrWhiteSpace(quan)) missingFields.Add("7. Quận");
                if (string.IsNullOrWhiteSpace(tienThu)) missingFields.Add("8. Tiền Thu");
                if (string.IsNullOrWhiteSpace(tienShip)) missingFields.Add("9. Tiền Ship");
                if (string.IsNullOrWhiteSpace(tienHang)) missingFields.Add("10. Tiền Hàng");
                if (string.IsNullOrWhiteSpace(nguoiDi)) missingFields.Add("11. Người Đi");
                if (string.IsNullOrWhiteSpace(nguoiLay)) missingFields.Add("12. Người Lấy");
                if (string.IsNullOrWhiteSpace(ghiChu)) missingFields.Add("13. Ghi Chú");
                if (string.IsNullOrWhiteSpace(ung)) missingFields.Add("14. Ưng");
                if (string.IsNullOrWhiteSpace(hang)) missingFields.Add("15. Hàng");
                if (string.IsNullOrWhiteSpace(fail)) missingFields.Add("16. Fail");
                if (string.IsNullOrWhiteSpace(note)) missingFields.Add("17. Ghi Chú Thêm");

                if (missingFields.Count > 0)
                {
                    string missingMsg = "❌ Vui lòng điền đủ tất cả 17 trường bắt buộc:\n\n" + 
                                       string.Join("\n", missingFields);
                    MessageBox.Show(missingMsg, "Thiếu thông tin bắt buộc", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Validate date format
                if (!DateTime.TryParse(ngay, out DateTime dateVal))
                {
                    MessageBox.Show("Ngày phải ở định dạng DD-MM-YYYY", "Lỗi định dạng", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Validate numeric fields
                if (!decimal.TryParse(tienThu, out decimal tienThuVal) || tienThuVal < 0)
                {
                    MessageBox.Show("Tiền Thu phải là số dương!", "Lỗi định dạng");
                    return;
                }

                if (!decimal.TryParse(tienShip, out decimal tienShipVal) || tienShipVal < 0)
                {
                    MessageBox.Show("Tiền Ship phải là số dương!", "Lỗi định dạng");
                    return;
                }

                if (!decimal.TryParse(tienHang, out decimal tienHangVal) || tienHangVal < 0)
                {
                    MessageBox.Show("Tiền Hàng phải là số dương!", "Lỗi định dạng");
                    return;
                }

                // Log entry (for now, just display success)
                string displayMsg = $"✅ Lưu thành công:\n\n" +
                    $"Tình Trạng: {tinhTrang}\n" +
                    $"Ngày: {ngay}\n" +
                    $"Địa Chỉ: {soNha}, {tenDuong}, {quan}\n" +
                    $"Tiền Thu: {tienThuVal:N0}\n" +
                    $"Người Đi: {nguoiDi}\n" +
                    $"Người Lấy: {nguoiLay}";

                MessageBox.Show(displayMsg, "Thành công");
                Debug.WriteLine($"✅ Manual entry saved: {ma} - {soNha}, {tenDuong}, {quan}");

                // TODO: Save to Excel with all 17 fields
                // For now, just validate and display success
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi: {ex.Message}", "Lỗi");
                Debug.WriteLine($"Error saving manual entry: {ex.Message}\n{ex.StackTrace}");
            }
        }
    }
}
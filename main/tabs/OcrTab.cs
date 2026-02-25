using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using Google.Cloud.Vision.V1;
using ClosedXML.Excel;
using TextInputter.Services;

namespace TextInputter
{
    /// <summary>
    /// OCR Tab: quét ảnh hàng loạt, hiển thị raw OCR log + mapping log, xuất Excel
    /// </summary>
    public partial class MainForm
    {
        // ─── Controls thuộc OCR Tab ────────────────────────────────────────────
        private TextBox txtNguoiDiOCR;
        private TextBox txtNguoiLayOCR;
        private RichTextBox txtRawOCRLog;
        private RichTextBox txtProcessLog;

        // ─── Init ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Khởi tạo tab OCR: folder selection, người đi/lấy, raw log, mapping log, export button
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
                UIHelper.CreateSectionLabel(pnlOCR, "🔍 OCR Processing", ref y);
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
                btnStartScan.Click += (s, e) => btnStart_Click(null, EventArgs.Empty);
                pnlOCR.Controls.Add(btnStartScan);

                y += 45;

                // ===== MANUAL INPUT SECTION: NGƯỜI ĐI & NGƯỜI LẤY =====
                UIHelper.CreateSectionLabel(pnlOCR, "Thong tin NGUOI DI & NGUOI LAY (bat buoc):", ref y);
                y -= 15;

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

                // ===== RAW OCR LOG =====
                UIHelper.CreateSectionLabel(pnlOCR, "📋 Raw OCR Text (Kết quả OCR thô):", ref y);
                y -= 15;

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

                // ===== MAPPING LOG =====
                UIHelper.CreateSectionLabel(pnlOCR, "✅ Chi tiet quet OCR (Mapping kết quả):", ref y);
                y -= 15;

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
                    { "rawLog",      this.txtRawOCRLog },
                    { "mappingLog",  this.txtProcessLog },
                    { "log",         batchLog },
                    { "checklist",   chkList }
                };

                // Responsive resize
                pnlOCR.Resize += (s, e) =>
                {
                    if (txtNguoiDiOCR  != null) txtNguoiDiOCR.Width  = pnlOCR.ClientSize.Width - 20;
                    if (txtNguoiLayOCR != null) txtNguoiLayOCR.Width = pnlOCR.ClientSize.Width - 20;
                    if (txtRawOCRLog   != null) txtRawOCRLog.Width   = pnlOCR.ClientSize.Width - 30;
                    if (txtProcessLog  != null) txtProcessLog.Width  = pnlOCR.ClientSize.Width - 30;
                };

                tabOCR.Controls.Clear();
                tabOCR.Controls.Add(pnlOCR);

                Debug.WriteLine("OCR Tab initialized");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error initializing OCR Tab: {ex.Message}");
            }
        }

        // ─── OCR Folder / Batch Processing ────────────────────────────────────

        /// <summary>
        /// Chọn folder chứa ảnh để batch OCR
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

                        // Cập nhật UI panel trái (giống hành vi cũ)
                        lblFolderPath.Text = folderPath;
                        lblImageCount.Text = $"{imageFiles.Count} ảnh";
                        lblStatus.Text     = $"✅ Đã chọn {imageFiles.Count} ảnh";
                        lblStatus.ForeColor = Color.Green;

                        // Cập nhật log box trong tab OCR nếu có
                        var pnlOCR = tabOCR.Controls[0] as Panel;
                        if (pnlOCR?.Tag is Dictionary<string, object> refs &&
                            refs.TryGetValue("log", out var logObj) &&
                            logObj is RichTextBox log)
                        {
                            log.Clear();
                            log.Text = $"📁 Folder: {folderPath}\n";
                            log.AppendText($"🖼️ Tìm thấy {imageFiles.Count} ảnh\n\nDanh sách:\n");
                            foreach (var img in imageFiles)
                                log.AppendText($"  • {Path.GetFileName(img)}\n");
                        }

                        Debug.WriteLine($"Selected folder: {folderPath}, Found {imageFiles.Count} images");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi");
                Debug.WriteLine($"Error selecting folder: {ex.Message}");
            }
        }

        /// <summary>
        /// Chạy batch OCR trên tất cả ảnh trong folder đã chọn
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

                var pnlOCR = tabOCR.Controls[0] as Panel;
                if (pnlOCR?.Tag is not Dictionary<string, object> refs) return;
                if (!refs.TryGetValue("log",       out var logObj)       || logObj       is not RichTextBox log)      return;
                if (!refs.TryGetValue("checklist", out var checkListObj) || checkListObj is not CheckedListBox chkList) return;

                log.Clear();
                log.Text = $"Quet {imageFiles.Count} anh...\n\n";

                int successCount = 0, failCount = 0;
                var failedImages    = new List<string>();
                var failedReasons   = new Dictionary<string, string>();
                var successImages   = new List<string>();

                chkList.Items.Clear();

                foreach (var imagePath in imageFiles)
                {
                    try
                    {
                        log.AppendText($"Xu ly: {Path.GetFileName(imagePath)}...\n");
                        Application.DoEvents();

                        var (ocrText, _) = CallPythonOCR(imagePath);
                        if (string.IsNullOrEmpty(ocrText))
                        {
                            log.AppendText("  [FAIL] OCR failed\n");
                            failCount++;
                            failedImages.Add(Path.GetFileName(imagePath));
                            failedReasons[Path.GetFileName(imagePath)] = "OCR text empty";
                            continue;
                        }

                        Dictionary<string, string> fields = new Dictionary<string, string>();
                        List<string> missingFields = new List<string>();

                        if (_ocrParsingService != null)
                            missingFields = _ocrParsingService.ExtractAllFields(ocrText, out fields) ?? new List<string>();

                        if (missingFields.Count > 0)
                        {
                            log.AppendText($"  [FAIL] Thieu: {string.Join(", ", missingFields)}\n");
                            failCount++;
                            failedImages.Add(Path.GetFileName(imagePath));
                            failedReasons[Path.GetFileName(imagePath)] = $"Missing: {string.Join(", ", missingFields)}";
                            continue;
                        }

                        string soHD = fields?.ContainsKey("Số HĐ") == true ? fields["Số HĐ"] : string.Empty;

                        if (_excelInvoiceService.InvoiceExists(soHD, out string existingSheet))
                        {
                            log.AppendText($"  [SKIP] SoHD '{soHD}' ton tai (sheet: {existingSheet})\n");
                            failCount++;
                            failedImages.Add(Path.GetFileName(imagePath));
                            continue;
                        }

                        string fileName = Path.GetFileName(imagePath);
                        chkList.Items.Add(fileName, true);
                        successImages.Add(imagePath);
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

                refs["successImages"] = successImages;

                log.AppendText($"\n{new string('=', 60)}\nKET QUA:\nOK: {successCount}/{imageFiles.Count}\nFAIL: {failCount}/{imageFiles.Count}\n");
                if (failedImages.Count > 0)
                {
                    log.AppendText("\nAnh that bai:\n");
                    foreach (var f in failedImages) log.AppendText($"  * {f}\n");
                }

                MessageBox.Show(
                    $"Hoan tat xu ly!\n\nThanh cong: {successCount}\nThat bai: {failCount}\n\nChon anh can xuat o duoi roi nhan 'Xuat'",
                    "Thong bao", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Loi: {ex.Message}", "Loi");
                Debug.WriteLine($"Error in batch processing: {ex.Message}\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Xuất các ảnh được tích chọn sang Excel.
        /// ⚠️ Hiện dùng _excelInvoiceService.ExportInvoice() — hardcoded path file Excel trong ExcelInvoiceService.
        /// </summary>
        private void ExportSelectedImages()
        {
            try
            {
                var pnlOCR = tabOCR.Controls[0] as Panel;
                if (pnlOCR?.Tag is not Dictionary<string, object> refs) return;
                if (!refs.TryGetValue("checklist",     out var checkListObj) || checkListObj is not CheckedListBox chkList)     return;
                if (!refs.TryGetValue("successImages",  out var successObj)  || successObj   is not List<string>   successImages) return;

                var selectedIndices = new List<int>();
                for (int i = 0; i < chkList.CheckedItems.Count; i++)
                    selectedIndices.Add(chkList.Items.IndexOf(chkList.CheckedItems[i]));

                if (selectedIndices.Count == 0)
                {
                    MessageBox.Show("Vui long chon it nhat 1 anh", "Canh bao");
                    return;
                }

                // NOTE: ExportSelectedImages chỉ dùng được khi _excelInvoiceService != null
                // (tức là file Excel mặc định tồn tại). Nếu muốn chọn file → dùng ExportMappedDataToExcel().
                MessageBox.Show("⚠️ Chức năng này yêu cầu file Excel cố định.\nDùng '💾 XUẤT EXCEL' bên dưới để chọn file.", "Thông báo");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi: {ex.Message}", "Lỗi");
                Debug.WriteLine($"Error exporting: {ex.Message}");
            }
        }

        // ─── Batch OCR → Map → Validate ───────────────────────────────────────

        /// <summary>
        /// Xử lý toàn bộ danh sách ảnh: OCR → Map → Validate → append vào mappedDataList.
        /// Chạy trên background thread (gọi từ btnStart_Click).
        /// </summary>
        private void ProcessImages()
        {
            var allText = new System.Text.StringBuilder();
            int successCount = 0, failCount = 0;
            mappedDataList.Clear();

            string nguoiDi  = txtNguoiDiOCR?.Text  ?? "";
            string nguoiLay = txtNguoiLayOCR?.Text ?? "";

            if (string.IsNullOrWhiteSpace(nguoiDi) || string.IsNullOrWhiteSpace(nguoiLay))
            {
                this.Invoke((MethodInvoker)delegate
                {
                    MessageBox.Show("❌ Vui lòng nhập NGƯỜI ĐI và NGƯỜI LẤY trước khi quét", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    btnStart.Enabled        = true;
                    btnSelectFolder.Enabled = true;
                    btnClear.Enabled        = true;
                    isProcessing = false;
                });
                return;
            }

            allText.AppendLine("╔════════════════════════════════════════════════════════╗");
            allText.AppendLine("║    KẾT QUẢ NHẬN DIỆN & MAP DỮ LIỆU (OCR) TIẾNG VIỆT   ║");
            allText.AppendLine("╚════════════════════════════════════════════════════════╝\n");
            allText.AppendLine($"📅 Ngày: {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
            allText.AppendLine($"📁 Folder: {folderPath}");
            allText.AppendLine($"👤 Người Đi: {nguoiDi} | Người Lấy: {nguoiLay}");
            allText.AppendLine($"📷 Tổng ảnh: {imageFiles.Count}");
            allText.AppendLine("\n" + new string('═', 60) + "\n");

            this.Invoke((MethodInvoker)delegate
            {
                txtResult.Text     = allText.ToString();
                txtProcessLog.Text = allText.ToString();
            });

            for (int i = 0; i < imageFiles.Count; i++)
            {
                string imagePath = imageFiles[i];
                string fileName  = Path.GetFileName(imagePath);

                this.Invoke((MethodInvoker)delegate
                {
                    progressBar.Value   = i + 1;
                    lblCurrentFile.Text = $"🔄 [{i + 1}/{imageFiles.Count}] {fileName}";
                });

                try
                {
                    var (text, confidence) = CallPythonOCR(imagePath);

                    // Header mỗi file — hiển thị ở CẢ HAI text area (có số thứ tự)
                    string fileHeader = $"\n{new string('═', 60)}\n📄 [{i + 1}/{imageFiles.Count}] {fileName}  (confidence: {confidence:F1}%)\n{new string('─', 60)}\n";

                    // Raw OCR log: chỉ raw text
                    this.Invoke((MethodInvoker)delegate
                    {
                        txtRawOCRLog?.AppendText(fileHeader + (text ?? "(Empty OCR result)") + "\n");
                    });

                    // Mapping log: chỉ hiển thị kết quả mapping (không lặp raw OCR)
                    allText.AppendLine(fileHeader);

                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        // Delegate field extraction to OCRTextParsingService
                        var missingFields = _ocrParsingService.ExtractAllFields(text, out var fields);

                        // Inject người đi/lấy from UI
                        fields["NGƯỜI ĐI"]  = nguoiDi;
                        fields["NGƯỜI LẤY"] = nguoiLay;

                        // Auto-fill TIỀN SHIP từ bảng phí ship theo quận (nếu chưa extract được)
                        if (string.IsNullOrWhiteSpace(fields.GetValueOrDefault("TIỀN SHIP", "")))
                        {
                            string quan = fields.GetValueOrDefault("QUẬN", "");
                            decimal? feeFromTable = OCRInvoiceMapper.GetShipFeeByQuan(quan);
                            if (feeFromTable.HasValue)
                            {
                                fields["TIỀN SHIP"] = feeFromTable.Value.ToString("F0");
                                allText.AppendLine($"  🗺️ Ship tự điền từ bảng: Q.{quan} → {feeFromTable.Value}k");
                            }
                            else
                            {
                                fields["TIỀN SHIP"] = "0";
                            }
                        }

                        // Compute TIỀN HÀNG = THU + SHIP
                        if (long.TryParse(fields.GetValueOrDefault("TIỀN THU",  ""), out long thu) &&
                            long.TryParse(fields.GetValueOrDefault("TIỀN SHIP", "0"), out long ship))
                            fields["TIỀN HÀNG"] = (thu + ship).ToString();

                        fields["fileName"] = fileName;

                        // Re-check missing after injecting manual fields
                        var stillMissing = missingFields.Where(f => string.IsNullOrWhiteSpace(fields.GetValueOrDefault(f, ""))).ToList();

                        if (stillMissing.Count == 0)
                        {
                            allText.AppendLine("📊 KẾT QUẢ MAP: ✅ THÀNH CÔNG — đủ fields");
                            foreach (var kv in fields.Where(k => k.Key != "fileName"))
                                allText.AppendLine($"  ✓ {kv.Key}: {kv.Value}");
                            mappedDataList.Add(fields);
                            successCount++;
                        }
                        else
                        {
                            allText.AppendLine($"📊 KẾT QUẢ MAP: ⚠️ THIẾU {stillMissing.Count} fields: {string.Join(", ", stillMissing)}");
                            // Log chi tiết từng field pass/fail
                            foreach (var kv in fields.Where(k => k.Key != "fileName"))
                            {
                                bool isMissing = stillMissing.Contains(kv.Key);
                                allText.AppendLine(isMissing
                                    ? $"  ✗ {kv.Key}: (trống)"
                                    : $"  ✓ {kv.Key}: {kv.Value}");
                            }
                            failCount++;
                        }
                    }
                    else
                    {
                        allText.AppendLine("📊 KẾT QUẢ MAP: ⚠️ Không nhận diện được text từ ảnh này");
                        failCount++;
                    }
                    // Không cần dòng kẻ cuối — header của file tiếp theo đã có kẻ ═══
                }
                catch (Exception ex)
                {
                    allText.AppendLine($"\n❌ TỆP #{i + 1}: {fileName} — Lỗi: {ex.Message}");
                    allText.AppendLine(new string('─', 60));
                    failCount++;
                }

                this.Invoke((MethodInvoker)delegate
                {
                    txtResult.Text               = allText.ToString();
                    txtResult.SelectionStart     = txtResult.Text.Length;
                    txtResult.ScrollToCaret();
                    txtProcessLog.Text           = allText.ToString();
                    txtProcessLog.SelectionStart = txtProcessLog.Text.Length;
                    txtProcessLog.ScrollToCaret();
                });
            }

            allText.AppendLine($"\n✅ Thành công: {successCount}/{imageFiles.Count}");
            allText.AppendLine($"❌ Thất bại:   {failCount}/{imageFiles.Count}");
            allText.AppendLine($"💾 Sẵn sàng xuất {mappedDataList.Count} dòng sang Excel");

            this.Invoke((MethodInvoker)delegate
            {
                txtResult.Text      = allText.ToString();
                txtProcessLog.Text  = allText.ToString();
                lblCurrentFile.Text = $"✅ Hoàn thành: {successCount} thành công, {failCount} thất bại";
                lblStatus.Text      = "✅ Xử lý xong";
                lblStatus.ForeColor = Color.Green;
                btnStart.Enabled        = true;
                btnSelectFolder.Enabled = true;
                btnClear.Enabled        = true;
                isProcessing = false;
                txtResult.SelectionStart = 0;
                txtResult.ScrollToCaret();

                // Lưu raw OCR log ra file
                string rawLog   = txtRawOCRLog?.Text ?? "";
                string logPath  = SaveOCRLog(rawLog);
                if (!string.IsNullOrEmpty(logPath))
                    lblCurrentFile.Text += $"  |  💾 Log: {logPath}";
            });
        }

        /// <summary>
        /// Ghi raw OCR log ra ocr_log.txt tại root project.
        /// File này nằm trong .gitignore — chỉ dùng để debug, không commit.
        /// </summary>
        private string SaveOCRLog(string content)
        {
            try
            {
                // BaseDirectory = bin/Debug/net8.0-windows → lên 3 cấp = root project
                string rootDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", ".."));
                string logPath = Path.Combine(rootDir, "ocr_log.txt");
                File.WriteAllText(logPath, content, System.Text.Encoding.UTF8);
                Debug.WriteLine($"✅ OCR log saved: {logPath}");
                return logPath;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"⚠️ Could not save OCR log: {ex.Message}");
                return "";
            }
        }

        // ─── Export Mapped Data → Excel ────────────────────────────────────────

        /// <summary>
        /// Xuất mappedDataList sang file Excel được chọn (user picks file, append vào sheet dd-MM).
        ///
        /// ⚠️ HARDCODED trong block này:
        ///   - Header array 20 columns — phụ thuộc format file Excel của khách.
        ///   - Sheet name = "dd-MM" lấy từ NGÀY LẤY của dòng đầu tiên.
        ///   - Row 2 ghi "THU x" / "NGAY x-x" theo cấu trúc file Excel mẫu.
        ///   - Data bắt đầu từ row 3.
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

                using var openDialog = new OpenFileDialog
                {
                    Filter           = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                    Title            = "Chọn file Excel để export dữ liệu",
                    InitialDirectory = Path.Combine(Directory.GetCurrentDirectory(), "data", "sample", "excel")
                };
                if (openDialog.ShowDialog() != DialogResult.OK) return;

                string excelPath = openDialog.FileName;
                if (!File.Exists(excelPath))
                {
                    MessageBox.Show($"❌ File không tồn tại: {excelPath}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Determine sheet name from data date
                var now = DateTime.Now;
                string sheetName = now.ToString("dd-MM");
                if (mappedDataList[0].TryGetValue("NGÀY LẤY", out string ngay) && !string.IsNullOrEmpty(ngay))
                {
                    var parts = ngay.Split('-');
                    if (parts.Length >= 2) sheetName = $"{parts[0]}-{parts[1]}";
                }

                DateTime sheetDate = now;
                DateTime.TryParseExact(sheetName, "dd-MM",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out sheetDate);

                // ⚠️ HARDCODED: 20-column header matching Excel template of current client
                var headers = new[]
                {
                    "Tình trạng TT", "SHOP", "TÊN KH", "MÃ", "SỐ NHÀ", "TÊN ĐƯỜNG", "QUẬN",
                    "TIỀN THU", "TIỀN SHIP", "TIỀN HÀNG",
                    "NGƯỜI ĐI", "NGƯỜI LẤY", "NGÀY LẤY", "GHI CHÚ",
                    "ỨNG TIỀN", "HÀNG TỒN", "FAIL", "Column1", "Column2", "Column3"
                };

                using (var workbook = new XLWorkbook(excelPath))
                {
                    bool isNewSheet = !workbook.TryGetWorksheet(sheetName, out var worksheet);
                    if (isNewSheet)
                    {
                        worksheet = workbook.Worksheets.Add(sheetName);
                        // Row 1: column headers
                        for (int col = 0; col < headers.Length; col++)
                        {
                            var cell = worksheet.Cell(1, col + 1);
                            cell.Value = headers[col];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                        }
                        // Row 2: THU x / NGAY x-x label
                        string thuText = sheetDate.DayOfWeek == DayOfWeek.Sunday
                            ? "CHU NHAT" : "THU " + ((int)sheetDate.DayOfWeek + 1);
                        worksheet.Cell(2, 2).Value = thuText;
                        worksheet.Cell(2, 2).Style.Font.Bold = true;
                        worksheet.Cell(2, 3).Value = $"NGAY {sheetDate.Day}-{sheetDate.Month}";
                        worksheet.Cell(2, 3).Style.Font.Bold = true;
                    }

                    // Data starts at row 3
                    int currentRow = 3;
                    var lastUsed = worksheet.LastRowUsed();
                    if (lastUsed != null && lastUsed.RowNumber() >= 3)
                        currentRow = lastUsed.RowNumber() + 1;

                    int addedCount = 0, updatedCount = 0;
                    foreach (var data in mappedDataList)
                    {
                        string ma = data.GetValueOrDefault("MÃ", "");

                        // Upsert: tìm row có MÃ trùng → ghi đè; không có → thêm dòng mới
                        int targetRow = -1;
                        if (!string.IsNullOrEmpty(ma))
                        {
                            foreach (var row in worksheet.RowsUsed())
                            {
                                if (row.RowNumber() <= 2) continue;
                                if (row.Cell(4).GetString() == ma) { targetRow = row.RowNumber(); break; }
                            }
                        }
                        bool isUpdate = targetRow > 0;
                        if (!isUpdate)
                        {
                            targetRow = currentRow;
                            currentRow++;
                        }

                        worksheet.Cell(targetRow,  1).Value = "";
                        worksheet.Cell(targetRow,  2).Value = data.GetValueOrDefault("SHOP",       "");
                        worksheet.Cell(targetRow,  3).Value = data.GetValueOrDefault("TÊN KH",     "");
                        worksheet.Cell(targetRow,  4).Value = ma;
                        worksheet.Cell(targetRow,  5).Value = data.GetValueOrDefault("SỐ NHÀ",     "");
                        worksheet.Cell(targetRow,  6).Value = data.GetValueOrDefault("TÊN ĐƯỜNG",  "");
                        worksheet.Cell(targetRow,  7).Value = data.GetValueOrDefault("QUẬN",       "");
                        worksheet.Cell(targetRow,  8).Value = data.GetValueOrDefault("TIỀN THU",   "");
                        worksheet.Cell(targetRow,  9).Value = data.GetValueOrDefault("TIỀN SHIP",  "");
                        worksheet.Cell(targetRow, 10).Value = data.GetValueOrDefault("TIỀN HÀNG",  "");
                        worksheet.Cell(targetRow, 11).Value = data.GetValueOrDefault("NGƯỜI ĐI",   "");
                        worksheet.Cell(targetRow, 12).Value = data.GetValueOrDefault("NGƯỜI LẤY",  "");
                        worksheet.Cell(targetRow, 13).Value = data.GetValueOrDefault("NGÀY LẤY",   "");

                        if (isUpdate) updatedCount++; else addedCount++;
                    }

                    workbook.SaveAs(excelPath);

                    this.Invoke((MethodInvoker)delegate
                    {
                        MessageBox.Show(
                            $"✅ Xuất thành công!\n\n➕ Thêm mới: {addedCount}\n✏️ Ghi đè: {updatedCount}\n📅 Sheet: {sheetName}\n📂 File: {Path.GetFileName(excelPath)}",
                            "✅ Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        lblStatus.Text      = $"✅ Xuất {addedCount} mới, {updatedCount} cập nhật → sheet '{sheetName}'";
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
    }
}

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

                var btnExport = UIHelper.CreateButton("Xuat", Color.Orange, 270, y, 80, 35);
                btnExport.Click += (s, e) => ExportSelectedImages();
                pnlOCR.Controls.Add(btnExport);

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

                        string ocrText = ExtractTextFromImage(imagePath);
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
        /// Xuất các ảnh được tích chọn sang Excel
        /// </summary>
        private void ExportSelectedImages()
        {
            try
            {
                var pnlOCR = tabOCR.Controls[0] as Panel;
                if (pnlOCR?.Tag is not Dictionary<string, object> refs) return;
                if (!refs.TryGetValue("checklist",    out var checkListObj) || checkListObj is not CheckedListBox chkList)     return;
                if (!refs.TryGetValue("successImages", out var successObj)  || successObj   is not List<string>   successImages) return;

                var selectedIndices = new List<int>();
                for (int i = 0; i < chkList.CheckedItems.Count; i++)
                    selectedIndices.Add(chkList.Items.IndexOf(chkList.CheckedItems[i]));

                if (selectedIndices.Count == 0)
                {
                    MessageBox.Show("Vui long chon it nhat 1 anh", "Canh bao");
                    return;
                }

                int exportCount = 0;
                foreach (int idx in selectedIndices)
                {
                    if (idx < 0 || idx >= successImages.Count) continue;
                    try
                    {
                        string ocrText = ExtractTextFromImage(successImages[idx]);
                        if (string.IsNullOrEmpty(ocrText)) continue;

                        string soHD    = _ocrParsingService?.ExtractInvoiceNumber(ocrText) ?? string.Empty;
                        string diaChi  = _ocrParsingService?.ExtractAddress(ocrText) ?? string.Empty;
                        decimal tongTien = _ocrParsingService?.ExtractTotalAmount(ocrText) ?? 0m;

                        if (string.IsNullOrEmpty(soHD) || string.IsNullOrEmpty(diaChi) || tongTien <= 0) continue;
                        if (_excelInvoiceService.InvoiceExists(soHD, out _)) continue;

                        decimal chietKhau = _ocrParsingService?.ExtractDiscount(ocrText) ?? 0m;
                        var invoice = new Services.OCRInvoiceData
                        {
                            SoHoaDon       = soHD,
                            DiaChi         = diaChi,
                            TongTienHang   = tongTien,
                            ChietKhau      = chietKhau,
                            TongThanhToan  = tongTien - chietKhau,
                            NguoiDi        = "OCR Auto",
                            NguoiLay       = "OCR Auto"
                        };
                        _excelInvoiceService.ExportInvoice(invoice);
                        exportCount++;
                    }
                    catch (Exception itemEx)
                    {
                        Debug.WriteLine($"Failed to export image: {itemEx.Message}");
                    }
                }

                MessageBox.Show(exportCount > 0
                    ? $"✅ Xuất thành công {exportCount} ảnh!"
                    : "⚠️ Không có ảnh nào được xuất thành công", "Thông báo");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi: {ex.Message}", "Lỗi");
                Debug.WriteLine($"Error exporting: {ex.Message}");
            }
        }

        /// <summary>
        /// Extract text from image (placeholder — hiện dùng Google Vision qua CallPythonOCR)
        /// </summary>
        private string ExtractTextFromImage(string imagePath)
        {
            try
            {
                return "";
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error extracting text: {ex.Message}");
                return "";
            }
        }

        // ─── OCR Data Processing (map, validate, export) ───────────────────────

        /// <summary>
        /// Xử lý toàn bộ danh sách ảnh: OCR → Map → Validate → append vào mappedDataList
        /// Được gọi từ btnStart_Click (tab cũ) hoặc StartBatchOCRProcessing
        /// </summary>
        private void ProcessImages()
        {
            System.Text.StringBuilder allText = new System.Text.StringBuilder();
            int successCount = 0, failCount = 0;
            mappedDataList.Clear();

            string nguoiDi  = txtNguoiDiOCR?.Text ?? "";
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
            allText.AppendLine($"👤 Người Đi: {nguoiDi}");
            allText.AppendLine($"👤 Người Lấy: {nguoiLay}");
            allText.AppendLine($"📷 Tổng ảnh: {imageFiles.Count}");
            allText.AppendLine("\n" + new string('═', 60) + "\n");

            this.Invoke((MethodInvoker)delegate
            {
                txtResult.Text       = allText.ToString();
                txtProcessLog.Text   = allText.ToString();
            });

            for (int i = 0; i < imageFiles.Count; i++)
            {
                string imagePath = imageFiles[i];
                string fileName  = Path.GetFileName(imagePath);

                this.Invoke((MethodInvoker)delegate
                {
                    progressBar.Value    = i + 1;
                    lblCurrentFile.Text  = $"🔄 [{i + 1}/{imageFiles.Count}] {fileName}";
                });

                try
                {
                    var (text, confidence) = CallPythonOCR(imagePath);

                    this.Invoke((MethodInvoker)delegate
                    {
                        if (txtRawOCRLog != null)
                        {
                            txtRawOCRLog.AppendText($"\n{new string('═', 60)}\n📄 TỆP: {fileName}\n📊 Độ tin cậy: {confidence:F1}%\n{new string('─', 60)}\n");
                            txtRawOCRLog.AppendText(text ?? "(Empty OCR result)\n");
                        }
                    });

                    allText.AppendLine($"\n✅ TỆP #{i + 1}: {fileName}");
                    allText.AppendLine($"   📊 Độ tin cậy: {confidence:F1}%");
                    allText.AppendLine($"   ⏱️  Thời gian: {DateTime.Now:HH:mm:ss}");
                    allText.AppendLine(new string('─', 60));

                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        var mappedData    = MapOCRDataTo12Fields(text, fileName, nguoiDi, nguoiLay);
                        var missingFields = ValidateMappedData(mappedData);
                        var fieldStatuses = GetFieldStatuses(mappedData);

                        if (missingFields.Count == 0)
                        {
                            allText.AppendLine("\n✅ THÀNH CÔNG - DỮ LIỆU ĐẦY ĐỦ (11/11 FIELDS):");
                            foreach (var key in new[] { "SHOP", "TÊN KH", "MÃ", "SỐ NHÀ", "TÊN ĐƯỜNG", "QUẬN",
                                                         "TIỀN THU", "TIỀN SHIP", "TIỀN HÀNG", "NGÀY LẤY", "NGƯỜI ĐI", "NGƯỜI LẤY" })
                                allText.AppendLine($"  ✓ {key}: {mappedData[key]}");

                            mappedDataList.Add(mappedData);
                            successCount++;
                        }
                        else
                        {
                            int passedCount = 11 - missingFields.Count;
                            allText.AppendLine($"\n⚠️ TỰA THÀNH CÔNG ({passedCount}/11 FIELDS):");
                            allText.AppendLine("   ✅ FIELDS PASS:");
                            foreach (var kvp in fieldStatuses)
                                if (kvp.Value) allText.AppendLine($"      ✓ {kvp.Key}: {mappedData[kvp.Key]}");
                            allText.AppendLine("   ❌ FIELDS FAIL:");
                            foreach (var field in missingFields)
                                allText.AppendLine($"      ✗ {field}");
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

                this.Invoke((MethodInvoker)delegate
                {
                    txtResult.Text            = allText.ToString();
                    txtResult.SelectionStart  = txtResult.Text.Length;
                    txtResult.ScrollToCaret();
                    txtProcessLog.Text        = allText.ToString();
                    txtProcessLog.SelectionStart = txtProcessLog.Text.Length;
                    txtProcessLog.ScrollToCaret();
                });
            }

            allText.AppendLine("\n\n╔════════════════════════════════════════════════════════╗");
            allText.AppendLine("║                    TÓM TẮT KẾT QUẢ                      ║");
            allText.AppendLine("╚════════════════════════════════════════════════════════╝\n");
            allText.AppendLine($"✅ Thành công: {successCount}/{imageFiles.Count}");
            allText.AppendLine($"❌ Thất bại:   {failCount}/{imageFiles.Count}");
            allText.AppendLine($"⏱️  Thời gian: {DateTime.Now:HH:mm:ss}");
            allText.AppendLine($"💾 Sẵn sàng xuất {mappedDataList.Count} dòng sang Excel");

            this.Invoke((MethodInvoker)delegate
            {
                txtResult.Text     = allText.ToString();
                txtProcessLog.Text = allText.ToString();
                lblCurrentFile.Text = $"✅ Hoàn thành: {successCount} thành công, {failCount} thất bại";
                lblStatus.Text      = "✅ Xử lý xong";
                lblStatus.ForeColor = Color.Green;
                btnStart.Enabled        = true;
                btnSelectFolder.Enabled = true;
                btnClear.Enabled        = true;
                isProcessing = false;
                txtResult.SelectionStart = 0;
                txtResult.ScrollToCaret();

                // Lưu raw OCR log ra file (ghi đè mỗi session)
                string rawLog = txtRawOCRLog?.Text ?? string.Empty;
                string savedPath = SaveOCRLog(rawLog);
                if (!string.IsNullOrEmpty(savedPath))
                    lblCurrentFile.Text += $"  |  💾 Log: {savedPath}";
            });
        }

        /// <summary>
        /// Lưu toàn bộ raw OCR log ra ocr_log.txt (ghi đè mỗi session).
        /// File nằm cùng folder với ảnh đã quét, hoặc thư mục app nếu chưa chọn folder.
        /// </summary>
        private string SaveOCRLog(string content)
        {
            try
            {
                // Lưu vào root project (cùng tầng .gitignore)
                // bin/Debug/net8.0-windows/ → lên 3 cấp = root project
                string rootDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", ".."));
                string logPath = Path.Combine(rootDir, "ocr_log.txt");

                File.WriteAllText(logPath, content, System.Text.Encoding.UTF8);
                Debug.WriteLine($"✅ OCR log saved: {logPath}");
                return logPath;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"⚠️ Could not save OCR log: {ex.Message}");
                return string.Empty;
            }
        }

        // ─── OCR Field Extraction Helpers ─────────────────────────────────────

        /// <summary>Map raw OCR text → 12 fields dictionary</summary>
        private Dictionary<string, string> MapOCRDataTo12Fields(string ocrText, string fileName, string nguoiDi, string nguoiLay)
        {
            var tienThu  = ExtractNumeric(ocrText, "tiền thu|thu tiền|tổng thanh toán");
            var tienShip = ExtractNumeric(ocrText, "tiền ship|ship|vận chuyển");
            if (string.IsNullOrEmpty(tienShip)) tienShip = "0";

            string tienHang = "";
            if (!string.IsNullOrEmpty(tienThu) || !string.IsNullOrEmpty(tienShip))
            {
                long thu  = long.TryParse(tienThu,  out var t) ? t : 0;
                long ship = long.TryParse(tienShip, out var s) ? s : 0;
                tienHang = (thu + ship).ToString();
            }

            string ngayLay = ExtractDateFromOCR(ocrText);
            if (string.IsNullOrEmpty(ngayLay))
                ngayLay = DateTime.Now.ToString("dd-MM-yyyy");

            return new Dictionary<string, string>
            {
                { "fileName",   fileName },
                { "SHOP",       ExtractField(ocrText, "đoàn|shop|cửa hàng", 100) },
                { "TÊN KH",     ExtractField(ocrText, "khách hàng:|customer:", 100) },
                { "NGƯỜI ĐI",   nguoiDi },
                { "NGƯỜI LẤY",  nguoiLay },
                { "MÃ",         ExtractField(ocrText, "so hd:|so hd|mã|ma:", 50) },
                { "SỐ NHÀ",     ExtractAddressField(ocrText, "soNha") },
                { "TÊN ĐƯỜNG",  ExtractAddressField(ocrText, "tenDuong") },
                { "QUẬN",       ExtractAddressField(ocrText, "quan") },
                { "TIỀN THU",   tienThu },
                { "TIỀN SHIP",  tienShip },
                { "TIỀN HÀNG",  tienHang },
                { "NGÀY LẤY",   ngayLay }
            };
        }

        /// <summary>
        /// Extract ngày tháng năm từ OCR.
        /// Hỗ trợ: "Ngày DD tháng MM năm YYYY" và "DD/MM/YYYY" / "DD-MM-YYYY"
        /// </summary>
        private string ExtractDateFromOCR(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";

            var m1 = System.Text.RegularExpressions.Regex.Match(text,
                @"ng[aà]y\s+(\d{1,2})\s+th[aá]ng\s+(\d{1,2})\s+n[aă]m\s+(\d{4})",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (m1.Success)
                return $"{m1.Groups[1].Value.PadLeft(2,'0')}-{m1.Groups[2].Value.PadLeft(2,'0')}-{m1.Groups[3].Value}";

            var m2 = System.Text.RegularExpressions.Regex.Match(text, @"\b(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})\b");
            if (m2.Success)
                return $"{m2.Groups[1].Value.PadLeft(2,'0')}-{m2.Groups[2].Value.PadLeft(2,'0')}-{m2.Groups[3].Value}";

            return "";
        }

        /// <summary>Extract address field từ OCR text (dùng AddressParser)</summary>
        private string ExtractAddressField(string ocrText, string fieldType)
        {
            if (string.IsNullOrWhiteSpace(ocrText)) return "";

            var lines = ocrText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            int addressBlockCount = 0, startLine = -1;

            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].IndexOf("địa chỉ", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    lines[i].IndexOf("địa chi", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    addressBlockCount++;
                    if (addressBlockCount == 2) { startLine = i; break; }
                }
            }

            if (startLine == -1)
            {
                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].IndexOf("địa chỉ", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        lines[i].IndexOf("địa chi", StringComparison.OrdinalIgnoreCase) >= 0)
                    { startLine = i; break; }
                }
            }

            if (startLine == -1) return "";

            string addressLine = lines[startLine];
            int colonIdx = addressLine.IndexOf(':');
            if (colonIdx >= 0) addressLine = addressLine.Substring(colonIdx + 1).Trim();

            var parsed = AddressParser.Parse(addressLine);
            return fieldType.ToLower() switch
            {
                "sonha"    => parsed.SoNha,
                "tenduong" => parsed.TenDuong,
                "quan"     => parsed.Quan,
                _          => addressLine
            };
        }

        /// <summary>Extract text field từ OCR text theo keyword</summary>
        private string ExtractField(string text, string keywords, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            var lines       = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var keywordList = keywords.Split('|');

            foreach (var line in lines)
            {
                foreach (var keyword in keywordList)
                {
                    if (line.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
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
        /// Extract số tiền từ OCR text.
        /// Trả về "" nếu không tìm thấy (không phải "0") để ValidateMappedData nhận biết thiếu.
        /// </summary>
        private string ExtractNumeric(string text, string keywords)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            var lines       = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var keywordList = keywords.Split('|');

            for (int i = 0; i < lines.Length; i++)
            {
                foreach (var keyword in keywordList)
                {
                    if (lines[i].IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Thử tìm số trên cùng dòng với keyword
                        var m = System.Text.RegularExpressions.Regex.Match(lines[i], @"[\d][,\d]*\d");
                        if (m.Success) return ToThousands(m.Value);

                        // Không có số → tìm ở dòng kế tiếp (pattern: "Tổng thanh toán:\n1,200,000")
                        if (i + 1 < lines.Length)
                        {
                            var next = System.Text.RegularExpressions.Regex.Match(lines[i + 1].Trim(), @"^[\d][,\d]*\d$");
                            if (next.Success) return ToThousands(next.Value);
                        }
                    }
                }
            }
            return "";
        }

        /// <summary>
        /// Chuyển số tiền dạng "1,200,000" hoặc "1200000" → đơn vị nghìn → "1200"
        /// Khớp với format Excel template (790 = 790,000 VND)
        /// </summary>
        private string ToThousands(string raw)
        {
            var digits = raw.Replace(",", "");
            if (long.TryParse(digits, out long val))
            {
                // Nếu số >= 1000 thì chia 1000 (đơn vị nghìn đồng)
                if (val >= 1000) return (val / 1000).ToString();
                return val.ToString();
            }
            return digits;
        }

        /// <summary>Validate mapped data — 11 required fields</summary>
        private List<string> ValidateMappedData(Dictionary<string, string> mappedData)
        {
            var required = new[] { "SHOP", "TÊN KH", "MÃ", "SỐ NHÀ", "TÊN ĐƯỜNG", "QUẬN",
                                   "TIỀN THU", "TIỀN SHIP", "NGÀY LẤY", "NGƯỜI ĐI", "NGƯỜI LẤY" };
            return required
                .Where(f => !mappedData.ContainsKey(f) || string.IsNullOrWhiteSpace(mappedData[f]))
                .ToList();
        }

        /// <summary>Get pass/fail status cho từng required field</summary>
        private Dictionary<string, bool> GetFieldStatuses(Dictionary<string, string> mappedData)
        {
            var required = new[] { "SHOP", "TÊN KH", "MÃ", "SỐ NHÀ", "TÊN ĐƯỜNG", "QUẬN",
                                   "TIỀN THU", "TIỀN SHIP", "NGÀY LẤY", "NGƯỜI ĐI", "NGƯỜI LẤY" };
            return required.ToDictionary(
                f => f,
                f => mappedData.ContainsKey(f) && !string.IsNullOrWhiteSpace(mappedData[f]));
        }

        // ─── Export Mapped Data → Excel ────────────────────────────────────────

        /// <summary>
        /// Xuất mappedDataList sang file Excel đã chọn (append vào sheet dd-MM)
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
                string sheetName;
                if (mappedDataList[0].ContainsKey("NGÀY LẤY") && !string.IsNullOrEmpty(mappedDataList[0]["NGÀY LẤY"]))
                {
                    var parts = mappedDataList[0]["NGÀY LẤY"].Split('-');
                    sheetName = parts.Length >= 2 ? $"{parts[0]}-{parts[1]}" : now.ToString("dd-MM");
                }
                else
                    sheetName = now.ToString("dd-MM");

                DateTime sheetDate = now;
                if (DateTime.TryParseExact(sheetName, "dd-MM",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var parsedDate))
                    sheetDate = parsedDate;

                using (var workbook = new XLWorkbook(excelPath))
                {
                    IXLWorksheet worksheet;
                    bool isNewSheet;

                    if (workbook.TryGetWorksheet(sheetName, out worksheet))
                    {
                        Debug.WriteLine($"✅ Sheet '{sheetName}' đã tồn tại, append dữ liệu");
                        isNewSheet = false;
                    }
                    else
                    {
                        worksheet  = workbook.Worksheets.Add(sheetName);
                        isNewSheet = true;
                        Debug.WriteLine($"✨ Tạo sheet mới: '{sheetName}'");
                    }

                    var headers = new[]
                    {
                        "Tình trạng TT", "SHOP", "TÊN KH", "MÃ", "SỐ NHÀ", "TÊN ĐƯỜNG", "QUẬN",
                        "TIỀN THU", "TIỀN SHIP", "TIỀN HÀNG",
                        "NGƯỜI ĐI", "NGƯỜI LẤY", "NGÀY LẤY", "GHI CHÚ",
                        "ỨNG TIỀN", "HÀNG TỒN", "FAIL", "Column1", "Column2", "Column3"
                    };

                    if (isNewSheet)
                    {
                        for (int col = 0; col < headers.Length; col++)
                        {
                            var cell = worksheet.Cell(1, col + 1);
                            cell.Value = headers[col];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                        }

                        string thuText  = sheetDate.DayOfWeek == DayOfWeek.Sunday ? "CHU NHAT" : "THU " + ((int)sheetDate.DayOfWeek + 1);
                        string ngayText = "NGAY " + sheetDate.Day + "-" + sheetDate.Month;

                        var cellThu  = worksheet.Cell(2, 2);
                        cellThu.Value = thuText;
                        cellThu.Style.Font.Bold = true;

                        var cellNgay = worksheet.Cell(2, 3);
                        cellNgay.Value = ngayText;
                        cellNgay.Style.Font.Bold = true;
                    }

                    int currentRow = 3;
                    var lastUsed   = worksheet.LastRowUsed();
                    if (lastUsed != null && lastUsed.RowNumber() >= 3)
                        currentRow = lastUsed.RowNumber() + 1;

                    int addedCount = 0;
                    foreach (var data in mappedDataList)
                    {
                        worksheet.Cell(currentRow, 1).Value  = "";
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
                        lblStatus.Text      = $"✅ Xuất {addedCount} dòng → sheet '{sheetName}'";
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

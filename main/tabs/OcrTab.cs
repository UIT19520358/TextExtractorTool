using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using TextInputter.Services;

namespace TextInputter
{
    /// <summary>
    /// OcrTab logic — SelectOCRFolder, ProcessImages, ExportMappedDataToExcel, SaveOCRLog.
    /// UI (control fields + InitializeOCRTab) ở OcrTab.UI.cs.
    /// </summary>
    public partial class MainForm
    {
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
                    var (text, confidence) = CallGoogleVisionOCR(imagePath);

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
        /// Logic ghi Excel được delegate sang <see cref="ExcelInvoiceService.ExportBatch"/>.
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

                var service = new ExcelInvoiceService(excelPath);
                var (addedCount, updatedCount) = service.ExportBatch(mappedDataList, sheetName, sheetDate);

                this.Invoke((MethodInvoker)delegate
                {
                    MessageBox.Show(
                        $"✅ Xuất thành công!\n\n➕ Thêm mới: {addedCount}\n✏️ Ghi đè: {updatedCount}\n📅 Sheet: {sheetName}\n📂 File: {Path.GetFileName(excelPath)}",
                        "✅ Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    lblStatus.Text      = $"✅ Xuất {addedCount} mới, {updatedCount} cập nhật → sheet '{sheetName}'";
                    lblStatus.ForeColor = Color.Green;
                });
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

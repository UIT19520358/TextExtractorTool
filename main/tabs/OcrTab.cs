using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
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
                        imageFiles = Directory
                            .GetFiles(folderPath, "*.*")
                            .Where(f =>
                                new[] { ".jpg", ".jpeg", ".png", ".bmp", ".tiff" }.Contains(
                                    Path.GetExtension(f).ToLower()
                                )
                            )
                            .OrderBy(f =>
                            {
                                // Sort tự nhiên: "1" < "2" < "10" (không phải "1" < "10" < "2")
                                var name = Path.GetFileNameWithoutExtension(f);
                                return int.TryParse(name, out int n) ? n : int.MaxValue;
                            })
                            .ThenBy(
                                f => Path.GetFileNameWithoutExtension(f),
                                StringComparer.OrdinalIgnoreCase
                            )
                            .ToList();

                        // Cập nhật UI panel trái (giống hành vi cũ)
                        lblFolderPath.Text = folderPath;
                        lblImageCount.Text = $"{imageFiles.Count} ảnh";
                        lblStatus.Text = $"✅ Đã chọn {imageFiles.Count} ảnh";
                        lblStatus.ForeColor = Color.Green;

                        Debug.WriteLine(
                            $"Selected folder: {folderPath}, Found {imageFiles.Count} images"
                        );
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
            var combinedLog = new System.Text.StringBuilder(); // unified per-image log (raw + mapping + Gemini)
            int successCount = 0,
                warnCount = 0;

            // Nếu đã có data từ batch trước, hỏi user muốn append hay replace
            if (mappedDataList.Count > 0)
            {
                var choice = MessageBox.Show(
                    $"Đã có {mappedDataList.Count} đơn từ lần quét trước.\n\n"
                        + "• YES = Giữ lại và cộng thêm batch mới\n"
                        + "• NO  = Xóa sạch và chỉ giữ batch mới",
                    "Giữ dữ liệu cũ?",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );
                if (choice == DialogResult.No)
                    mappedDataList.Clear();
                // Nếu YES → giữ nguyên, append tiếp bên dưới
            }

            string nguoiDi = txtNguoiDiOCR?.Text ?? "";
            string nguoiLay = txtNguoiLayOCR?.Text ?? "";

            // nguoiDi/nguoiLay có thể để trống vì sẽ auto-map theo khu vực từng đơn

            allText.AppendLine("╔════════════════════════════════════════════════════════╗");
            allText.AppendLine("║    KẾT QUẢ NHẬN DIỆN & MAP DỮ LIỆU (OCR) TIẾNG VIỆT   ║");
            allText.AppendLine("╚════════════════════════════════════════════════════════╝\n");
            allText.AppendLine($"📅 Ngày: {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
            allText.AppendLine($"📁 Folder: {folderPath}");
            allText.AppendLine($"👤 Người Đi: {nguoiDi} | Người Lấy: {nguoiLay}");
            allText.AppendLine($"📷 Tổng ảnh: {imageFiles.Count}");
            allText.AppendLine("\n" + new string('═', 60) + "\n");

            combinedLog.AppendLine($"=== OCR RUN: {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
            combinedLog.AppendLine($"📁 Folder: {folderPath}");
            combinedLog.AppendLine($"👤 Người Đi: {nguoiDi} | Người Lấy: {nguoiLay}");
            combinedLog.AppendLine($"📷 Tổng ảnh: {imageFiles.Count}");
            combinedLog.AppendLine();

            this.Invoke(
                (MethodInvoker)
                    delegate
                    {
                        txtResult.Text = allText.ToString();
                        txtProcessLog.Text = allText.ToString();
                    }
            );

            for (int i = 0; i < imageFiles.Count; i++)
            {
                string imagePath = imageFiles[i];
                string fileName = Path.GetFileName(imagePath);

                this.Invoke(
                    (MethodInvoker)
                        delegate
                        {
                            progressBar.Value = i + 1;
                            lblCurrentFile.Text = $"🔄 [{i + 1}/{imageFiles.Count}] {fileName}";
                        }
                );

                try
                {
                    var (text, confidence) = CallGoogleVisionOCR(imagePath);

                    // Header mỗi file — hiển thị ở CẢ HAI text area (có số thứ tự)
                    string fileHeader =
                        $"\n{new string('═', 60)}\n📄 [{i + 1}/{imageFiles.Count}] {fileName}  (confidence: {confidence:F1}%)\n{new string('─', 60)}\n";

                    // Raw OCR log: chỉ raw text
                    this.Invoke(
                        (MethodInvoker)
                            delegate
                            {
                                txtRawOCRLog?.AppendText(
                                    fileHeader + (text ?? "(Empty OCR result)") + "\n"
                                );
                            }
                    );

                    // Mapping log: chỉ hiển thị kết quả mapping (không lặp raw OCR)
                    allText.AppendLine(fileHeader);

                    // combinedLog — ghi header + raw OCR trước
                    combinedLog.AppendLine(new string('═', 60));
                    combinedLog.AppendLine(
                        $"📄 [{i + 1}/{imageFiles.Count}] {fileName}  (confidence: {confidence:F1}%)"
                    );
                    combinedLog.AppendLine(new string('─', 60));
                    combinedLog.AppendLine("[RAW OCR]");
                    combinedLog.AppendLine(text ?? "(Empty OCR result)");
                    combinedLog.AppendLine();

                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        // Set image path để GeminiService fallback biết đọc ảnh nào
                        _ocrParsingService.CurrentImagePath = imagePath;

                        // List để nhận log Gemini — sau đó gom vào combinedLog
                        var geminiLog = new List<string>();

                        // Delegate field extraction to OCRTextParsingService
                        var missingFields = _ocrParsingService.ExtractAllFields(
                            text,
                            out var fields,
                            geminiLog
                        );

                        // Auto-map người đi theo phường/quận, hoặc dùng giá trị tự nhập
                        string phuongForMap = fields.GetValueOrDefault("PHƯỜNG", "");
                        string quanForMap = fields.GetValueOrDefault("QUẬN", "");
                        fields["NGƯỜI ĐI"] =
                            _manualNguoiDi && !string.IsNullOrWhiteSpace(txtNguoiDiOCR?.Text)
                                ? txtNguoiDiOCR.Text.Trim()
                                : OCRInvoiceMapper.GetNguoiDi(phuongForMap, quanForMap);
                        fields["NGƯỜI LẤY"] =
                            _manualNguoiLay && !string.IsNullOrWhiteSpace(txtNguoiLayOCR?.Text)
                                ? txtNguoiLayOCR.Text.Trim()
                                : nguoiLay;

                        // Auto-fill TIỀN SHIP từ bảng phí ship theo phường/quận (tier-3 → tier-2)
                        // Điều kiện auto-fill: TIỀN SHIP rỗng HOẶC = "0" (tức là chưa biết thực sự)
                        string currentShip = fields.GetValueOrDefault("TIỀN SHIP", "");
                        bool shipIsUnknown =
                            string.IsNullOrWhiteSpace(currentShip) || currentShip.Trim() == "0";
                        if (shipIsUnknown)
                        {
                            string phuong = fields.GetValueOrDefault("PHƯỜNG", "");
                            string quan = fields.GetValueOrDefault("QUẬN", "");
                            decimal? feeFromTable = OCRInvoiceMapper.GetShipFee(phuong, quan);
                            if (feeFromTable.HasValue)
                            {
                                fields["TIỀN SHIP"] = feeFromTable.Value.ToString("F0");
                                allText.AppendLine(
                                    $"  🗺️ Ship tự điền từ bảng: Q.{quan} P.{phuong} → {feeFromTable.Value}k"
                                );
                            }
                            else
                            {
                                // Không có trong bảng → để trống, user tự điền sau
                                // (KHÔNG gán "0" vì "0" != rỗng sẽ block auto-fill lần sau)
                                fields["TIỀN SHIP"] = "";
                                allText.AppendLine(
                                    $"  ⚠️ Ship chưa có trong bảng: Q.{quan} P.{phuong} — cần điền tay"
                                );
                            }
                        }

                        // Compute TIỀN HÀNG theo loại đơn:
                        //   COD          : thu + ship  (format cũ)
                        //   SHIP_ONLY_FREE: -ship       (không thu ship, tiền hàng âm)
                        //   SHIP_ONLY_PAID: +ship       (thu ship, tiền hàng = ship)
                        string invoiceType = fields.GetValueOrDefault("INVOICE_TYPE", "COD");
                        long.TryParse(fields.GetValueOrDefault("TIỀN THU", "0"), out long thu);
                        long.TryParse(fields.GetValueOrDefault("TIỀN SHIP", "0"), out long ship);
                        long tienhang = invoiceType switch
                        {
                            "SHIP_ONLY_FREE" => -ship,
                            "SHIP_ONLY_PAID" => ship,
                            _ => thu + ship, // COD
                        };
                        fields["TIỀN HÀNG"] = tienhang.ToString();
                        // Log loại đơn ra UI nếu không phải COD
                        if (invoiceType != "COD")
                            allText.AppendLine(
                                $"  📦 Loại đơn: {invoiceType} → TIỀN HÀNG = {tienhang}"
                            );

                        // Log TÌNH TRẠNG đặc biệt (đã CK, hàng sỉ, v.v.)
                        string tinhTrangDetected = fields.GetValueOrDefault("TÌNH TRẠNG", "");
                        if (!string.IsNullOrEmpty(tinhTrangDetected))
                            allText.AppendLine($"  🏷️  Tình trạng: {tinhTrangDetected}");

                        fields["fileName"] = fileName;

                        // Re-check missing after injecting manual fields
                        var stillMissing = missingFields
                            .Where(f => string.IsNullOrWhiteSpace(fields.GetValueOrDefault(f, "")))
                            .ToList();

                        // Ghi Gemini log vào combinedLog (nếu có)
                        if (geminiLog.Count > 0)
                        {
                            combinedLog.AppendLine("[GEMINI]");
                            foreach (var gLine in geminiLog)
                                combinedLog.AppendLine(gLine);
                            combinedLog.AppendLine();
                        }

                        string mappingResult;
                        if (stillMissing.Count == 0)
                        {
                            mappingResult = "📊 KẾT QUẢ MAP: ✅ THÀNH CÔNG — đủ fields";
                            allText.AppendLine(mappingResult);
                            combinedLog.AppendLine("[MAPPING]");
                            combinedLog.AppendLine(mappingResult);
                            foreach (var kv in fields.Where(k => k.Key != "fileName"))
                            {
                                string line = $"  ✓ {kv.Key}: {kv.Value}";
                                allText.AppendLine(line);
                                combinedLog.AppendLine(line);
                            }
                            fields["IS_FAIL"] = "0";
                            fields["MISSING_FIELDS"] = "";
                            successCount++;
                        }
                        else
                        {
                            // Thiếu field → vẫn ghi Excel bình thường, chỉ tô đỏ cell bị thiếu
                            mappingResult =
                                $"📊 KẾT QUẢ MAP: ⚠️ THIẾU {stillMissing.Count} fields: {string.Join(", ", stillMissing)} — đã lưu, cần check thủ công";
                            allText.AppendLine(mappingResult);
                            combinedLog.AppendLine("[MAPPING]");
                            combinedLog.AppendLine(mappingResult);
                            foreach (var kv in fields.Where(k => k.Key != "fileName"))
                            {
                                bool isMissing = stillMissing.Contains(kv.Key);
                                string line = isMissing
                                    ? $"  ⚠️ {kv.Key}: (trống)"
                                    : $"  ✓ {kv.Key}: {kv.Value}";
                                allText.AppendLine(line);
                                combinedLog.AppendLine(line);
                            }
                            fields["IS_FAIL"] = "0"; // không đánh dấu fail — tính bình thường
                            fields["MISSING_FIELDS"] = string.Join(",", stillMissing); // để Excel tô đỏ từng cell
                            warnCount++;
                        }
                        mappedDataList.Add(fields); // cả đủ field lẫn thiếu field đều lưu
                        combinedLog.AppendLine();
                    }
                    else
                    {
                        string noText = "📊 KẾT QUẢ MAP: ⚠️ Không nhận diện được text từ ảnh này";
                        allText.AppendLine(noText);
                        combinedLog.AppendLine("[MAPPING]");
                        combinedLog.AppendLine(noText);
                        combinedLog.AppendLine();
                        // Đơn không OCR được vẫn lưu vào Excel (để trống, tô đỏ GHI CHÚ)
                        var emptyFields = new Dictionary<string, string>
                        {
                            ["fileName"] = fileName,
                            ["IS_FAIL"] = "0",
                            ["MISSING_FIELDS"] = "GHI CHÚ",
                            ["GHI CHÚ"] = $"OCR thất bại: {fileName}",
                        };
                        mappedDataList.Add(emptyFields);
                        warnCount++;
                    }
                    // Không cần dòng kẻ cuối — header của file tiếp theo đã có kẻ ═══
                }
                catch (Exception ex)
                {
                    allText.AppendLine($"\n❌ TỆP #{i + 1}: {fileName} — Lỗi: {ex.Message}");
                    allText.AppendLine(new string('─', 60));
                    combinedLog.AppendLine($"[ERROR] {fileName}: {ex.Message}");
                    combinedLog.AppendLine();
                    var errFields = new Dictionary<string, string>
                    {
                        ["fileName"] = fileName,
                        ["IS_FAIL"] = "0",
                        ["MISSING_FIELDS"] = "GHI CHÚ",
                        ["GHI CHÚ"] = $"Lỗi: {ex.Message}",
                    };
                    mappedDataList.Add(errFields);
                    warnCount++;
                }

                this.Invoke(
                    (MethodInvoker)
                        delegate
                        {
                            txtResult.Text = allText.ToString();
                            txtResult.SelectionStart = txtResult.Text.Length;
                            txtResult.ScrollToCaret();
                            txtProcessLog.Text = allText.ToString();
                            txtProcessLog.SelectionStart = txtProcessLog.Text.Length;
                            txtProcessLog.ScrollToCaret();
                        }
                );
            }

            // Tất cả đơn (đủ field + thiếu field) đã được add vào mappedDataList trong vòng lặp
            // (đơn thiếu field có MISSING_FIELDS != "" → Excel tô đỏ từng cell tương ứng)

            allText.AppendLine($"\n✅ Đủ fields:      {successCount}/{imageFiles.Count}");
            if (warnCount > 0)
                allText.AppendLine(
                    $"⚠️ Thiếu field:   {warnCount}/{imageFiles.Count} (đã lưu, cần điền tay)"
                );
            allText.AppendLine($"💾 Sẵn sàng xuất {mappedDataList.Count} dòng sang Excel");

            combinedLog.AppendLine(new string('═', 60));
            combinedLog.AppendLine($"✅ Đủ fields: {successCount}/{imageFiles.Count}");
            if (warnCount > 0)
                combinedLog.AppendLine($"⚠️ Thiếu field: {warnCount}/{imageFiles.Count}");

            this.Invoke(
                (MethodInvoker)
                    delegate
                    {
                        txtResult.Text = allText.ToString();
                        txtProcessLog.Text = allText.ToString();
                        lblCurrentFile.Text =
                            warnCount > 0
                                ? $"✅ Hoàn thành: {successCount} đủ fields, {warnCount} cần check"
                                : $"✅ Hoàn thành: {successCount} đơn";
                        lblStatus.Text = "✅ Xử lý xong";
                        lblStatus.ForeColor = Color.Green;
                        btnStart.Enabled = true;
                        btnSelectFolder.Enabled = true;
                        btnClear.Enabled = true;
                        isProcessing = false;
                        txtResult.SelectionStart = 0;
                        txtResult.ScrollToCaret();

                        // Ghi combined log (raw OCR + mapping + Gemini per-image) ra file
                        string logPath = SaveCombinedLog(combinedLog.ToString());
                        if (!string.IsNullOrEmpty(logPath))
                            lblCurrentFile.Text += $"  |  💾 Log: {logPath}";
                    }
            );
        }

        /// <summary>
        /// Ghi unified log (raw OCR + mapping + Gemini, per-image) ra ocr_log.txt tại root project.
        /// File này nằm trong .gitignore — chỉ dùng để debug, không commit.
        /// </summary>
        private string SaveCombinedLog(string content)
        {
            try
            {
                // BaseDirectory = bin/Debug/net8.0-windows → lên 3 cấp = root project
                string rootDir = Path.GetFullPath(
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..")
                );
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

        // ─── Load From Log ─────────────────────────────────────────────────────

        /// <summary>
        /// Đọc ocr_log.txt từ lần quét trước → khôi phục mappedDataList mà không tốn quota API.
        /// </summary>
        private void LoadFromLog()
        {
            try
            {
                // Tìm file log: ưu tiên cạnh .exe, fallback root project
                string rootDir = Path.GetFullPath(
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..")
                );
                string defaultLogPath = Path.Combine(rootDir, "ocr_log.txt");

                using var dlg = new OpenFileDialog
                {
                    Title = "Chọn file log OCR (ocr_log.txt)",
                    Filter = "Log files (*.txt)|*.txt|All files (*.*)|*.*",
                    FileName = "ocr_log.txt",
                    InitialDirectory = File.Exists(defaultLogPath)
                        ? rootDir
                        : AppDomain.CurrentDomain.BaseDirectory,
                };
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                string logPath = dlg.FileName;
                string[] lines = File.ReadAllLines(logPath, System.Text.Encoding.UTF8);

                // Nếu đã có data cũ → hỏi
                if (mappedDataList.Count > 0)
                {
                    var choice = MessageBox.Show(
                        $"Đã có {mappedDataList.Count} đơn trong bộ nhớ.\n\n"
                            + "• YES = Giữ lại và cộng thêm từ log\n"
                            + "• NO  = Xóa sạch và chỉ dùng log",
                        "Giữ dữ liệu cũ?",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question
                    );
                    if (choice == DialogResult.No)
                        mappedDataList.Clear();
                }

                // ── Parse log ─────────────────────────────────────────────────
                var wantedFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "SHOP",
                    "TÊN KH",
                    "MÃ",
                    "QUẬN",
                    "PHƯỜNG",
                    "TÊN ĐƯỜNG",
                    "ĐỊA CHỈ",
                    "TIỀN THU",
                    "TIỀN SHIP",
                    "TIỀN HÀNG",
                    "INVOICE_TYPE",
                    "TÌNH TRẠNG",
                    "NGÀY LẤY",
                    "NGƯỜI ĐI",
                    "NGƯỜI LẤY",
                    "GHI CHÚ",
                    "IS_FAIL",
                    "MISSING_FIELDS",
                };

                int loaded = 0;
                Dictionary<string, string> currentFields = null;
                string currentFileName = null;
                bool inMapping = false;
                bool isMissingBlock = false;
                var missingList = new List<string>();

                void FlushCurrent()
                {
                    if (currentFields == null)
                        return;
                    currentFields["fileName"] = currentFileName ?? "";
                    if (!currentFields.ContainsKey("MISSING_FIELDS"))
                        currentFields["MISSING_FIELDS"] =
                            isMissingBlock && missingList.Count > 0
                                ? string.Join(",", missingList)
                                : "";
                    if (!currentFields.ContainsKey("IS_FAIL"))
                        currentFields["IS_FAIL"] = "0";
                    mappedDataList.Add(currentFields);
                    loaded++;
                    currentFields = null;
                    currentFileName = null;
                    inMapping = false;
                    isMissingBlock = false;
                    missingList.Clear();
                }

                foreach (string rawLine in lines)
                {
                    string line = rawLine.TrimEnd();

                    // Bắt đầu block mới: dòng có "📄 [i/n] filename"
                    if (line.Contains("📄 [") && line.Contains("]"))
                    {
                        FlushCurrent();
                        int bracket = line.IndexOf(']');
                        if (bracket >= 0 && bracket + 1 < line.Length)
                        {
                            string afterBracket = line.Substring(bracket + 1).Trim();
                            int paren = afterBracket.IndexOf('(');
                            currentFileName =
                                paren > 0
                                    ? afterBracket.Substring(0, paren).Trim()
                                    : afterBracket.Trim();
                        }
                        inMapping = false;
                        isMissingBlock = false;
                        missingList.Clear();
                        continue;
                    }

                    // Bắt đầu section [MAPPING]
                    if (line == "[MAPPING]")
                    {
                        inMapping = true;
                        currentFields = new Dictionary<string, string>(
                            StringComparer.OrdinalIgnoreCase
                        );
                        continue;
                    }

                    if (!inMapping)
                        continue;

                    // Dòng kết quả tổng
                    if (line.Contains("KẾT QUẢ MAP"))
                    {
                        isMissingBlock = line.Contains("THIẾU") || line.Contains("⚠");
                        continue;
                    }

                    // Dòng field đủ: "  ✓ FIELD: value" hoặc "  ✔ FIELD: value"
                    if ((line.Contains("✓ ") || line.Contains("✔ ")) && line.Contains(":"))
                    {
                        // Tìm vị trí tick mark (✓ U+2713 hoặc ✔ U+2714)
                        int tick = line.IndexOf('✓');
                        if (tick < 0)
                            tick = line.IndexOf('✔');
                        int colon = line.IndexOf(':', tick);
                        if (tick >= 0 && colon > tick)
                        {
                            string key = line.Substring(tick + 1, colon - tick - 1).Trim();
                            string val =
                                colon + 1 < line.Length ? line.Substring(colon + 1).Trim() : "";
                            if (wantedFields.Contains(key))
                                currentFields[key] = val;
                        }
                        continue;
                    }

                    // Dòng field thiếu: "  ⚠ FIELD: (trống)"
                    if (line.Contains("⚠ ") && line.Contains(": (trống)"))
                    {
                        int warn = line.IndexOf('⚠');
                        int colon = line.IndexOf(':', warn);
                        if (warn >= 0 && colon > warn)
                        {
                            string key = line.Substring(warn + 1, colon - warn - 1).Trim();
                            if (wantedFields.Contains(key))
                            {
                                currentFields[key] = "";
                                missingList.Add(key);
                            }
                        }
                        continue;
                    }
                }

                FlushCurrent(); // flush block cuối

                // ── Update UI ─────────────────────────────────────────────────
                string summary =
                    $"✅ Tải từ log: {loaded} đơn\n"
                    + $"📄 File: {Path.GetFileName(logPath)}\n"
                    + $"🚀 Sẵn sàng xuất {mappedDataList.Count} dòng sang Excel";

                if (txtProcessLog != null)
                    txtProcessLog.Text = summary;
                txtRawOCRLog?.Clear();
                txtRawOCRLog?.AppendText($"[Tải từ log — không quét OCR]\n{summary}");
                lblStatus.Text = $"✅ Log: {loaded} đơn";
                lblStatus.ForeColor = Color.Green;
                lblCurrentFile.Text = $"📄 {loaded} đơn từ {Path.GetFileName(logPath)}";

                MessageBox.Show(
                    summary,
                    "✅ Tải từ log thành công",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌ Lỗi đọc log:\n{ex.Message}",
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
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
                    MessageBox.Show(
                        "❌ Không có dữ liệu để xuất. Vui lòng quét ảnh trước!",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }

                using var openDialog = new OpenFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                    Title = "Chọn file Excel để export dữ liệu",
                    InitialDirectory = Path.Combine(
                        Directory.GetCurrentDirectory(),
                        "data",
                        "sample",
                        "excel"
                    ),
                };
                if (openDialog.ShowDialog() != DialogResult.OK)
                    return;

                string excelPath = openDialog.FileName;
                if (!File.Exists(excelPath))
                {
                    MessageBox.Show(
                        $"❌ File không tồn tại: {excelPath}",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                    return;
                }

                // ── Group by sheet name theo mode user chọn ───────────────────
                var now = DateTime.Now;

                // Deduplicate theo MÃ trước khi export: giữ entry CUỐI CÙNG cho mỗi MÃ
                // (trường hợp user quét 2 folder append → cùng MÃ xuất hiện 2 lần)
                // Đơn MÃ rỗng (isMissing) giữ tất cả vì không thể so sánh
                var deduped = new List<Dictionary<string, string>>();
                var seenMa = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                // Duyệt ngược để giữ entry cuối (mới nhất) khi có trùng
                for (int di = mappedDataList.Count - 1; di >= 0; di--)
                {
                    var d = mappedDataList[di];
                    string ma = d.GetValueOrDefault("MÃ", "");
                    if (string.IsNullOrWhiteSpace(ma))
                    {
                        deduped.Insert(0, d); // MÃ rỗng → giữ tất cả, đặt theo thứ tự gốc
                    }
                    else if (seenMa.Add(ma))
                    {
                        deduped.Insert(0, d); // MÃ chưa thấy → giữ, đặt trước (giữ thứ tự gốc)
                    }
                    // Nếu đã thấy MÃ này → bỏ qua (entry sau = entry mới hơn đã được giữ)
                }
                int dupCount = mappedDataList.Count - deduped.Count;

                // Validate ngày tự nhập nếu đang ở mode "Ngày khác"
                if (_exportUseToday == null)
                {
                    // Phải có format dd-MM
                    bool validCustom = false;
                    if (!string.IsNullOrWhiteSpace(_exportCustomDate))
                    {
                        var p = _exportCustomDate.Trim().Split('-');
                        validCustom =
                            p.Length == 2
                            && int.TryParse(p[0], out int dd)
                            && dd >= 1
                            && dd <= 31
                            && int.TryParse(p[1], out int mm)
                            && mm >= 1
                            && mm <= 12;
                    }
                    if (!validCustom)
                    {
                        MessageBox.Show(
                            "Vui lòng nhập ngày hợp lệ theo định dạng  dd-MM\nVD: 25-02",
                            "Ngày không hợp lệ",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );
                        return;
                    }
                }

                // Hàm lấy sheetName "dd-MM" từ 1 dict
                string GetSheetName(Dictionary<string, string> d)
                {
                    // Mode: ngày tự nhập → tất cả vào sheet ngày đó
                    if (_exportUseToday == null)
                        return _exportCustomDate.Trim();

                    // Mode: ngày hôm nay → tất cả vào 1 sheet
                    if (_exportUseToday == true)
                        return now.ToString("dd-MM");

                    // Mode: theo ngày hóa đơn
                    if (d.TryGetValue("NGÀY LẤY", out string ngay) && !string.IsNullOrEmpty(ngay))
                    {
                        // Format gốc: "27-02-2026." hoặc "27-02" hoặc "27-02-2026"
                        var parts = ngay.TrimEnd('.').Split('-');
                        if (parts.Length >= 2)
                            return $"{parts[0]}-{parts[1]}";
                    }
                    return now.ToString("dd-MM");
                }

                var grouped = deduped.GroupBy(d => GetSheetName(d)).OrderBy(g => g.Key).ToList();

                var service = new ExcelInvoiceService(excelPath);
                int totalAdded = 0,
                    totalUpdated = 0;
                var sheetSummaries = new List<string>();

                foreach (var group in grouped)
                {
                    string sheetName = group.Key;
                    DateTime sheetDate = now;
                    DateTime.TryParseExact(
                        sheetName,
                        "dd-MM",
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None,
                        out sheetDate
                    );
                    // Gán năm hiện tại vì TryParseExact không có năm
                    if (sheetDate.Year == 1)
                        sheetDate = sheetDate.AddYears(now.Year - 1);

                    var (addedCount, updatedCount) = service.ExportBatch(
                        group.ToList(),
                        sheetName,
                        sheetDate
                    );
                    totalAdded += addedCount;
                    totalUpdated += updatedCount;
                    sheetSummaries.Add(
                        $"  📅 Sheet [{sheetName}]: +{addedCount} mới, ✏️{updatedCount} ghi đè ({group.Count()} ảnh)"
                    );
                }

                string detailText = string.Join("\n", sheetSummaries);
                string dupNote =
                    dupCount > 0
                        ? $"\n⚠️ Đã bỏ {dupCount} đơn trùng MÃ (giữ lần quét mới nhất)"
                        : "";

                this.Invoke(
                    (MethodInvoker)
                        delegate
                        {
                            MessageBox.Show(
                                $"✅ Xuất thành công!\n\n{detailText}{dupNote}\n\n➕ Tổng thêm mới: {totalAdded}\n✏️ Tổng ghi đè: {totalUpdated}\n📂 File: {Path.GetFileName(excelPath)}",
                                "✅ Thành công",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information
                            );
                            lblStatus.Text =
                                $"✅ Xuất {totalAdded} mới, {totalUpdated} cập nhật → {grouped.Count} sheet";
                            lblStatus.ForeColor = Color.Green;
                        }
                );
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ LỖI: {ex.Message}\n{ex.StackTrace}");
                this.Invoke(
                    (MethodInvoker)
                        delegate
                        {
                            MessageBox.Show(
                                $"❌ Lỗi xuất Excel:\n\n{ex.Message}",
                                "Lỗi",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                        }
                );
            }
        }
    }
}

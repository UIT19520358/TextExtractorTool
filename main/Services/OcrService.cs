using System;
using System.Collections.Generic;
using System.Linq;
using Google.Cloud.Vision.V1;

namespace TextInputter.Services
{
    /// <summary>
    /// Xử lý OCR: gọi Google Vision API, clean text rác.
    /// </summary>
    public class OcrService
    {
        private readonly ImageAnnotatorClient _visionClient;

        public OcrService(ImageAnnotatorClient visionClient)
        {
            _visionClient = visionClient;
        }

        /// <summary>
        /// Gọi Google Vision API để nhận diện text từ ảnh
        /// </summary>
        public (string text, float confidence) ExtractTextFromImage(string imagePath)
        {
            try
            {
                if (_visionClient == null)
                {
                    System.Diagnostics.Debug.WriteLine("ERROR: visionClient is null");
                    return ("", 0);
                }

                var image = Google.Cloud.Vision.V1.Image.FromFile(imagePath);
                var response = _visionClient.DetectTextAsync(image);
                response.Wait();

                if (response.Result == null || response.Result.Count == 0)
                {
                    return ("", 0);
                }

                var textAnnotation = response.Result[0];
                if (textAnnotation == null)
                {
                    return ("", 0);
                }

                string text = textAnnotation.Description?.Trim() ?? "";

                if (string.IsNullOrEmpty(text))
                {
                    return ("", 0);
                }

                // Lọc text rác
                text = CleanOCRText(text);

                if (string.IsNullOrEmpty(text))
                {
                    return ("", 0);
                }

                float confidence = 95.0f;
                return (text, confidence);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Google Vision error: {ex.Message}");
                return ("", 0);
            }
        }

        /// <summary>
        /// Lọc text rác từ kết quả OCR
        /// </summary>
        public string CleanOCRText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var cleanLines = new List<string>();

            foreach (var line in lines)
            {
                string trimmed = line.Trim();

                if (string.IsNullOrWhiteSpace(trimmed))
                    continue;

                if (IsGarbageLine(trimmed))
                    continue;

                if (trimmed.Length < 3)
                    continue;

                cleanLines.Add(trimmed);
            }

            return string.Join("\n", cleanLines);
        }

        /// <summary>
        /// Kiểm tra dòng có phải text rác không
        /// </summary>
        private bool IsGarbageLine(string line)
        {
            int validCharCount = 0;
            int totalCharCount = 0;

            foreach (char c in line)
            {
                totalCharCount++;

                bool isVietnamese = (c >= '\u0100' && c <= '\u01FF') ||
                                   (c >= '\u1E00' && c <= '\u1EFF');

                bool isEnglish = char.IsLetterOrDigit(c) ||
                                char.IsWhiteSpace(c) ||
                                c == ',' || c == '.' || c == '-' ||
                                c == '/' || c == ':' || c == ';' ||
                                c == '(' || c == ')';

                if (isVietnamese || isEnglish)
                    validCharCount++;
            }

            return validCharCount < (totalCharCount * 0.7);
        }
    }
}

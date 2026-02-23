using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace TextInputter.Services
{
    /// <summary>
    /// Xử lý File operations
    /// </summary>
    public class FileService
    {
        /// <summary>
        /// Lấy tất cả image files từ folder
        /// </summary>
        public List<string> GetImageFiles(string folderPath)
        {
            var imageFiles = new List<string>();

            try
            {
                if (!Directory.Exists(folderPath))
                    return imageFiles;

                var extensions = new[] { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff" };
                var files = Directory.GetFiles(folderPath);

                imageFiles = files
                    .Where(f => extensions.Contains(Path.GetExtension(f).ToLower()))
                    .ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi quét folder:\n{ex.Message}", "Lỗi");
            }

            return imageFiles;
        }

        /// <summary>
        /// Lưu text vào file
        /// </summary>
        public bool SaveTextToFile(string text, string fileName = "")
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                    DefaultExt = "txt",
                    FileName = fileName
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(saveFileDialog.FileName, text);
                    MessageBox.Show($"✅ Lưu thành công:\n{saveFileDialog.FileName}", "Thành công");
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi lưu file:\n{ex.Message}", "Lỗi");
            }

            return false;
        }

        /// <summary>
        /// Mở file dialog để chọn file
        /// </summary>
        public string OpenFileDialog(string filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*")
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = filter,
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileName;
            }

            return null;
        }

        /// <summary>
        /// Mở folder dialog
        /// </summary>
        public string OpenFolderDialog()
        {
            var folderBrowserDialog = new FolderBrowserDialog
            {
                Description = "Chọn thư mục",
                ShowNewFolderButton = false
            };

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                return folderBrowserDialog.SelectedPath;
            }

            return null;
        }

        /// <summary>
        /// Kiểm tra file có tồn tại không
        /// </summary>
        public bool FileExists(string filePath)
        {
            return File.Exists(filePath);
        }

        /// <summary>
        /// Xóa file
        /// </summary>
        public bool DeleteFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi xóa file:\n{ex.Message}", "Lỗi");
            }

            return false;
        }

        /// <summary>
        /// Lấy tên file từ path
        /// </summary>
        public string GetFileName(string filePath)
        {
            return Path.GetFileName(filePath);
        }

        /// <summary>
        /// Lấy folder path từ full path
        /// </summary>
        public string GetDirectoryPath(string filePath)
        {
            return Path.GetDirectoryName(filePath);
        }

        /// <summary>
        /// Combine paths
        /// </summary>
        public string CombinePath(params string[] paths)
        {
            return Path.Combine(paths);
        }

        /// <summary>
        /// Lấy application base directory
        /// </summary>
        public string GetAppBaseDirectory()
        {
            return AppDomain.CurrentDomain.BaseDirectory;
        }
    }
}

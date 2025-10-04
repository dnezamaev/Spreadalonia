using Avalonia.Controls;
using Avalonia.Platform.Storage;

using OfficeOpenXml;

using System.IO;
using System.Threading.Tasks;

namespace Demo
{
    internal class FileHelpers
    {
        public static void InitEpplus()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public static async Task<string?> ShowOpenFileDialog(Window window)
        {
            string? filePath = null;

            var topLevel = TopLevel.GetTopLevel(window);

            var files =
                await topLevel
                .StorageProvider
                .OpenFilePickerAsync(
                    new FilePickerOpenOptions
                    {
                        AllowMultiple = false,
                        FileTypeFilter = new[]
                        {
                            new FilePickerFileType("Excel files")
                            {
                                Patterns = new[] { "*.xlsx" },
                            },
                        }
                    });

            if (files.Count == 1)
            {
                filePath = files[0].Path.AbsolutePath;
            }

            return filePath;
        }

        public static ExcelPackage? LoadExcelPackage(string? filePath)
        {
            if (filePath is null)
            {
                return null;
            }

            var fileContent = File.ReadAllBytes(filePath);
            var package = new ExcelPackage(new MemoryStream(fileContent));
            return package;
        }

        public static ExcelPackage CreateEmptyExcelPackage()
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Sheet1");
            worksheet.Cells["A1"].Value = "Hello!";
            return package;
        }
    }
}

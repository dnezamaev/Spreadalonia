using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Platform.Storage;

using OfficeOpenXml;

using System.IO;
using System.Threading.Tasks;

namespace Demo
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            LoadDocument();
        }

        private void LoadDocument()
        {
            FileHelpers.InitEpplus();

            ExcelPackage excel;

            if (Design.IsDesignMode)
            {
                // Keep Avalonia designer simple and fast.
                // Load clean Excel document.
                excel = FileHelpers.CreateEmptyExcelPackage();
            }
            else
            {
                // Load sample content-rich Excel document for application.
                excel = FileHelpers.LoadExcelPackage("sample_workbook.xlsx");
            }

            spreadbook.LoadEpplusDocument(excel.Workbook);
        }

        private async void OpenFileMenuItem_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            var filePath = await FileHelpers.ShowOpenFileDialog(this);
            var epplus = FileHelpers.LoadExcelPackage(filePath);

            if (epplus is null)
            {
                return;
            }

            spreadbook.LoadEpplusDocument(epplus.Workbook);
        }

        private void Spreadbook_AllSpreadsheetsInitialized(object? sender, System.EventArgs e)
        {
        }
    }
}
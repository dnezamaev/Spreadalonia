using Avalonia.Controls;

using OfficeOpenXml;

namespace Demo
{
    public partial class MainWindow : Window
    {
        ExcelPackage excel;

        public MainWindow()
        {
            InitializeComponent();

            LoadDocument();
        }

        private void LoadDocument()
        {
            FileHelpers.InitEpplus();

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
            excel = FileHelpers.LoadExcelPackage(filePath);

            if (excel is null)
            {
                return;
            }

            spreadbook.LoadEpplusDocument(excel.Workbook);
        }

        private async void SaveFileMenuItem_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            var filePath = await FileHelpers.ShowSaveFileDialog(this);

            // Reload excel package, because it is closed after save.
            excel = FileHelpers.SaveExcelPackage(excel, filePath);
            spreadbook.LoadEpplusDocument(excel.Workbook);
        }

        private void Spreadbook_AllSpreadsheetsLoaded(object? sender, System.EventArgs e)
        {
        }
    }
}
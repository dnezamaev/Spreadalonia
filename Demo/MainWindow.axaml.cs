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

            FileHelpers.InitEpplus();

            // Load clean Excel document.
            var cleanExcel = FileHelpers.CreateEmptyExcelPackage();
            //spreadbook.LoadEpplusDocument(cleanExcel.Workbook);

            // Load a sample Excel document.
            spreadbook.LoadEpplusDocument(FileHelpers.LoadExcelPackage("sample_workbook.xlsx").Workbook);
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
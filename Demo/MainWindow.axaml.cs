using System;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using Spreadalonia;

namespace Demo;

public partial class MainWindow : Window
{
    private ExcelPackage excel;

    private DatePickerWindow datePickerWindow;

    private SpreadCell clickedCell;

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
            excel = FileHelpers.LoadExcelPackage(@"SampleWorkbooks\3_sheets_5_rows.xlsx");
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

    private void InsertRowMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        var spreadsheet = spreadbook.SelectedSpreadsheet as SpreadsheetEpplus;
        spreadsheet.InsertRowWithFormulas();
    }

    private void AppendRowMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        var spreadsheet = spreadbook.SelectedSpreadsheet as SpreadsheetEpplus;
        spreadsheet.AppendRowWithFormulas();
    }

    private void CopyRowMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        var spreadsheet = spreadbook.SelectedSpreadsheet as SpreadsheetEpplus;
        var selections = spreadsheet.Selection;

        if (selections.Count != 1) return;

        var selection = selections.First();

        spreadsheet.CopyRow(selection.Top, selection.Bottom);
    }

    /// <summary>
    /// Show calendar on double click cell with date.
    /// </summary>
    private async void Spreadbook_OnCellDoubleClicked(object? sender, SpreadCell cell)
    {
        var spreadsheet = cell.Worksheet as SpreadsheetEpplus;
        clickedCell = cell;

        // Apply only for first column of first worksheet.
        if (!(spreadsheet.WorksheetIndex == 0 && cell.Column == 0)) return;

        var currentCellValue = spreadsheet.GetCellValue(cell.Column, cell.Row);

        var startDate = currentCellValue switch
        {
            double doubleCellValue => DateTime.FromOADate(doubleCellValue),
            DateTime dateTimeCellValue => dateTimeCellValue,
            _ => DateTime.Now
        };

        // For example restrict date to 01 of Jan, Jul or Oct of any year.
        var options = new DatePickerWindowOptions
        {
            StartDate = startDate,
            DateFormat = "dd.MM.yyyy",
            RegexValidationPattern = @"01\.(01|07|10)\.\d{4}",
        };

        datePickerWindow = new DatePickerWindow();
        var vm = new DatePickerWindowViewModel(options); 
        datePickerWindow.DataContext = vm;
        vm.DatePicked += DatePickerWindowViewModel_DatePicked;
        await datePickerWindow.ShowDialog(this);
    }

    private void DatePickerWindowViewModel_DatePicked(object? sender, DateTime selectedDate)
    {
        datePickerWindow.Close();

        var spreadsheet = clickedCell.Worksheet as SpreadsheetEpplus;
        spreadsheet.SetCellValue(clickedCell.Column, clickedCell.Row, selectedDate);
    }
}
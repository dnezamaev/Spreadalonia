using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Avalonia.Controls;
using OfficeOpenXml;

namespace Spreadalonia;

public class SpreadbookEpplus : Spreadbook
{
    public ExcelWorkbook EpplusWorkbook { get; private set; }

    /// <summary>
    /// Load excel document.
    /// </summary>
    public void LoadEpplusDocument(ExcelWorkbook workbook)
    {
        this.EpplusWorkbook = 
            workbook ?? throw new ArgumentNullException(nameof(workbook));

        spreadsheets.Clear();
        bookTabControl.Items.Clear();

        var epplusWorksheets = GetEpplusWorksheets(workbook);

        EpplusWorkbook.Calculate();
        FillTabControl(epplusWorksheets);
    }

    private static IEnumerable<ExcelWorksheet> GetEpplusWorksheets(ExcelWorkbook workbook)
    {
        return workbook.Worksheets
            .Where(s => s.Hidden == eWorkSheetHidden.Visible);
    }

    private void FillTabControl(IEnumerable<ExcelWorksheet> epplusWorksheets)
    {
        foreach (var worksheet in epplusWorksheets)
        {
            AddTabItem(worksheet);
        }

        // Delay setting data, because Spreadsheet.Data is null until Spreadsheet is initialised.
        foreach (var spreadsheet in spreadsheets)
        {
            spreadsheet.Loaded += OnSpreadsheetLoaded;
        }

        // Force spreadsheets initialization.
        foreach (TabItem tabItem in SpreadsheetsTabControl.Items)
        {
            tabItem.IsSelected = true;
        }

        // Show user first tab.
        (SpreadsheetsTabControl.Items.First() as TabItem).IsSelected = true;
    }

    private void AddTabItem(ExcelWorksheet worksheet)
    {
        var spreadsheet = new SpreadsheetEpplus
        {
            SpreadBook = this,
            Options = WorksheetOptions,
            EpplusWorksheet = worksheet
        };

        spreadsheets.Add(spreadsheet);

        var tabItem = new TabItem
        {
            Content = spreadsheet,
            Header = worksheet.Name
        };

        bookTabControl.Items.Add(tabItem);
    }

    void OnSpreadsheetLoaded(object? o, EventArgs eventArgs)
    {
        var spreadsheet = o as SpreadsheetEpplus;
        var worksheet = spreadsheet.EpplusWorksheet;
        Debug.WriteLine($"OnSpreadsheetLoaded begins for {worksheet.Name}");

        // This handler must be called once.
        spreadsheet.Loaded -= OnSpreadsheetLoaded;

        var data = spreadsheet.GetDataFromEpplus();
        spreadsheet.SetSkipValues(true);
        spreadsheet.SetData(data);
        spreadsheet.SetSkipValues(false);

        var columnWidth = GetColumnsWidth(worksheet);
        spreadsheet.SetWidth(columnWidth);

        spreadsheet.AutoFitHeightAllRows();

        var hiddenRows = GetHiddenRows(worksheet);
        var hiddenColumns = GetHiddenColumns(worksheet);

        spreadsheet.SetHeight(hiddenRows);
        spreadsheet.SetWidth(hiddenColumns);

        if (spreadsheets.All(s => s.IsLoaded))
        {
            (bookTabControl.Items.FirstOrDefault() as TabItem).IsSelected = true;
            OnAllSpreadsheetsLoaded();
        }
    }

    private static Dictionary<int, double> GetHiddenRows(ExcelWorksheet worksheet)
    {
        var hiddenRows = new Dictionary<int, double>();

        for (int i = 1; i < worksheet.Dimension?.End?.Row; i++)
        {
            var row = worksheet.Rows[i];

            if (!row.Hidden)
            {
                continue;
            }

            hiddenRows[i - 1] = 0;
        }

        return hiddenRows;
    }

    private static Dictionary<int, double> GetHiddenColumns(ExcelWorksheet worksheet)
    {
        var hiddenColumns = new Dictionary<int, double>();

        for (int i = 1; i < worksheet.Dimension?.End?.Column; i++)
        {
            var column = worksheet.Columns[i];

            if (!column.Hidden)
            {
                continue;
            }

            hiddenColumns[i - 1] = 0;
        }

        return hiddenColumns;
    }

    private static Dictionary<int, double> GetColumnsWidth(ExcelWorksheet worksheet)
    {
        var widthDictionary = new Dictionary<int, double>();

        for (int i = 1; i < worksheet.Dimension?.End?.Column; i++)
        {
            var column = worksheet.Columns[i];

            widthDictionary[i - 1] = column.Width * 7.5;
        }

        return widthDictionary;
    }

    internal void ReloadDataFromEpplus()
    {
        EpplusWorkbook.Calculate();

        foreach (SpreadsheetEpplus spreadsheet in Spreadsheets)
        {
            var data = spreadsheet.GetDataFromEpplus();

            spreadsheet.SetSkipValues(true);
            spreadsheet.InnerClearAllContents();
            spreadsheet.SetData(data);
            spreadsheet.SetSkipValues(false);
        }
    }

}
using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using OfficeOpenXml;

namespace Spreadalonia;

/// <summary>
/// Excel-like set of spreadsheets.
/// </summary>
public partial class Spreadbook : UserControl
{
    #region Properties, fields.

    /// <summary>
    /// Contains one spreadsheet per tab.
    /// </summary>
    private readonly TabControl bookTabControl;

    private readonly List<Spreadsheet> spreadsheets;

    /// <summary>
    /// All spreadsheets of this book.
    /// </summary>
    public IEnumerable<Spreadsheet> Spreadsheets => spreadsheets;

    /// <summary>
    /// Defines the <see cref="WorksheetOptions"/> property.
    /// </summary>
    public static readonly StyledProperty<SpreadsheetOptions> WorksheetOptionsProperty = AvaloniaProperty.Register<Spreadsheet, SpreadsheetOptions>(nameof(WorksheetOptions), new SpreadsheetOptions());

    /// <summary>
    /// Allow cut cell content?
    /// </summary>
    public SpreadsheetOptions WorksheetOptions
    {
        get => GetValue(WorksheetOptionsProperty);
        set => SetValue(WorksheetOptionsProperty, value);
    }

    /// <summary>
    /// Active worksheet.
    /// </summary>
    public Spreadsheet SelectedSpreadsheet => 
        bookTabControl.SelectedContent as Spreadsheet;

    public ExcelWorkbook EpplusWorkbook { get; private set; }

    #endregion

    #region Events.

    /// <summary>
    /// Raises when all worksheets are initialized.
    /// </summary>
    public event EventHandler AllSpreadsheetsInitialized;

    public event EventHandler<MassiveSpreadCellsEditStartedEventArgs> MassiveCellsEditStarted;

    internal MassiveSpreadCellsEditStartedEventArgs RaiseMassiveCellsEditStarted
        (MassiveSpreadCellsEditStartedEventArgs eventArgs)
    {
        MassiveCellsEditStarted?.Invoke(this, eventArgs);
        return eventArgs;
    }

    public event EventHandler<SpreadCellEditStartedEventArgs> CellEditStarted;

    internal SpreadCellEditStartedEventArgs RaiseCellEditStarted(SpreadCellEditStartedEventArgs eventArgs)
    {
        CellEditStarted?.Invoke(this, eventArgs);
        return eventArgs;
    }

    public event EventHandler<SpreadCellEditEndedEventArgs> CellEditEnded;

    internal SpreadCellEditEndedEventArgs RaiseCellEditEnded(SpreadCellEditEndedEventArgs eventArgs)
    {
        CellEditEnded?.Invoke(this, eventArgs);

        if (eventArgs.Cancel) return eventArgs;

        UpdateEpplusCellValue(eventArgs.Cell, eventArgs.NewText);
        EpplusWorkbook.Calculate();
        ReloadDataFromEpplus();

        return eventArgs;
    }

    private void UpdateEpplusCellValue(SpreadCell cell, string newText)
    {
        var worksheetName = cell.Worksheet.WorksheetName;
        var epplusWorksheet = FindSpreadsheetInfo(worksheetName).EpplusWorksheet;

        epplusWorksheet.Cells[cell.Row + 1, cell.Column + 1].Value = newText;
    }

    private void ReloadDataFromEpplus()
    {
        foreach (var epplusWorksheet in EpplusWorkbook.Worksheets)
        {
            var data = GetDataForSpreadsheet(epplusWorksheet);
            var spreadsheet = FindSpreadsheet(epplusWorksheet.Name);
            spreadsheet.SetData(data);
        }
    }

    public event EventHandler<SpreadCell> CellDoubleClicked;

    internal void RaiseCellDoubleClicked(SpreadCell cell)
    {
        CellDoubleClicked?.Invoke(this, cell);
    }

    public event EventHandler<SpreadCell> CellClicked;

    internal void RaiseCellClicked(SpreadCell cell)
    {
        CellClicked?.Invoke(this, cell);
    }

    #endregion

    #region Initialization.

    public Spreadbook()
    {
        InitializeComponent();

        bookTabControl = this.FindControl<TabControl>(nameof(SpreadsheetsTabControl));
        spreadsheets = new List<Spreadsheet>();
    }

    #endregion

    #region Getting spreadsheets and their info.

    public Spreadsheet FindSpreadsheet(string sheetName)
    {
        return spreadsheets.FirstOrDefault(s => s.WorksheetName == sheetName);
    }

    public Spreadsheet FindSpreadsheet(int sheetIndex)
    {
        return spreadsheets.FirstOrDefault(s => s.WorksheetIndex == sheetIndex);
    }

    public SpreadsheetInfo FindSpreadsheetInfo(string sheetName)
    {
        return spreadsheets.FirstOrDefault(s => s.WorksheetName == sheetName)?.Tag as SpreadsheetInfo;
    }

    public SpreadsheetInfo FindSpreadsheetInfo(int sheetIndex)
    {
        return spreadsheets.FirstOrDefault(s => s.WorksheetIndex == sheetIndex)?.Tag as SpreadsheetInfo;
    }

    #endregion

    #region Loading Excel document.

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
    }

    private static IEnumerable<KeyValuePair<(int, int), string>> GetDataForSpreadsheet(ExcelWorksheet worksheet)
    {
        var dataForSpreadsheet = worksheet.Cells
            .Where(c => c.Value != null && !c.EntireRow.Hidden && !c.EntireColumn.Hidden)
            .Select(EpplusCellToSpreadsheetCell);

        return dataForSpreadsheet;
    }

    private static KeyValuePair<(int, int), string> EpplusCellToSpreadsheetCell(ExcelRangeBase c)
    {
        var spreadsheetCells = new KeyValuePair<(int, int), string>(
            key: new ValueTuple<int, int>(c.EntireColumn.StartColumn - 1, c.EntireRow.StartRow - 1),
            value: c.Value.ToString());

        return spreadsheetCells;
    }

    private void AddTabItem(ExcelWorksheet worksheet)
    {
        var spreadsheet = new Spreadsheet
        {
            Container = this,
            WorksheetIndex = worksheet.Index,
            WorksheetName = worksheet.Name,
            Options = WorksheetOptions
        };

        spreadsheets.Add(spreadsheet);

        spreadsheet.Initialized += OnSpreadsheetInitialized;

        var tabItem = new TabItem
        {
            Content = spreadsheet,
            Header = worksheet.Name
        };

        bookTabControl.Items.Add(tabItem);

        var tag = new SpreadsheetInfo
        {
            TabItem = tabItem,
            Spreadsheet = spreadsheet,
            EpplusWorksheet = worksheet
        };

        spreadsheet.Tag = tag;

        // Force spreadsheets initialization.
        tabItem.IsSelected = true;
    }

    void OnSpreadsheetInitialized(object? o, EventArgs eventArgs)
    {
        var spreadsheet = o as Spreadsheet;
        var tag = spreadsheet.Tag as SpreadsheetInfo;
        var worksheet = tag.EpplusWorksheet;

        var data = GetDataForSpreadsheet(worksheet);
        spreadsheet.SetData(data);

        var columnWidth = GetColumnsWidth(worksheet);
        spreadsheet.SetWidth(columnWidth);

        spreadsheet.AutoFitHeightAllRows();

        var hiddenRows = GetHiddenRows(worksheet);
        var hiddenColumns = GetHiddenColumns(worksheet);

        spreadsheet.SetHeight(hiddenRows);
        spreadsheet.SetWidth(hiddenColumns);

        if (spreadsheets.All(s => s.IsInitialized))
        {
            (bookTabControl.Items.FirstOrDefault() as TabItem).IsSelected = true;
            OnAllSpreadsheetsInitialized();
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

    protected virtual void OnAllSpreadsheetsInitialized()
    {
        AllSpreadsheetsInitialized?.Invoke(this, EventArgs.Empty);
    }

    #endregion
}

public class SpreadsheetInfo
{
    public TabItem TabItem { get; set; }
    public Spreadsheet Spreadsheet { get; set; }
    public ExcelWorksheet EpplusWorksheet { get; set; }
}
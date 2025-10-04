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
        return eventArgs;
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

    public Spreadbook()
    {
        InitializeComponent();

        bookTabControl = this.FindControl<TabControl>(nameof(SpreadsheetsTabControl));
        spreadsheets = new List<Spreadsheet>();
    }

    /// <summary>
    /// Load excel document.
    /// </summary>
    public void LoadEpplusDocument(ExcelWorkbook workbook)
    {
        spreadsheets.Clear();

        var eppWorksheets = GetEppWorksheets(workbook);

        FillTabControl(eppWorksheets);
    }

    public Spreadsheet FindSpreadsheet(string sheetName)
    {
        return spreadsheets.FirstOrDefault(s => s.WorksheetName == sheetName);
    }

    public Spreadsheet FindSpreadsheet(int sheetIndex)
    {
        return spreadsheets.FirstOrDefault(s => s.WorksheetIndex == sheetIndex);
    }

    private static IEnumerable<ExcelWorksheet> GetEppWorksheets(ExcelWorkbook workbook)
    {
        return workbook.Worksheets
            .Where(s => s.Hidden == eWorkSheetHidden.Visible);
    }

    private void FillTabControl(IEnumerable<ExcelWorksheet> eppWorksheets)
    {
        foreach (var worksheet in eppWorksheets)
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
        var tag = new SpreadsheetInfo
        {
            EppWorksheet = worksheet
        };

        var spreadsheet = new Spreadsheet
        {
            Tag = tag,
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
    }

    void OnSpreadsheetInitialized(object? o, EventArgs eventArgs)
    {
        var spreadsheet = o as Spreadsheet;
        var tag = spreadsheet.Tag as SpreadsheetInfo;
        var worksheet = tag.EppWorksheet;

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
            OnAllSpreadsheetsInitialized();
        }
    }

    private static Dictionary<int, double> GetHiddenRows(ExcelWorksheet worksheet)
    {
        var hiddenRows = new Dictionary<int, double>();

        for (int i = 1; i < worksheet.Dimension.End.Row; i++)
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

        for (int i = 1; i < worksheet.Dimension.End.Column; i++)
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

        for (int i = 1; i < worksheet.Dimension.End.Column; i++)
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
}

public class SpreadsheetInfo
{
    public ExcelWorksheet EppWorksheet { get; set; }
}
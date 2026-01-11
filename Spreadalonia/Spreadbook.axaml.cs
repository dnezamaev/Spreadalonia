using System;
using System.Collections.Generic;
using System.Diagnostics;
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
    protected readonly TabControl bookTabControl;

    protected readonly List<Spreadsheet> spreadsheets;

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

    #region Events & handlers.

    /// <summary>
    /// Raises when all worksheets are initialized.
    /// </summary>
    public event EventHandler AllSpreadsheetsLoaded;

    protected virtual void OnAllSpreadsheetsLoaded()
    {
        Debug.WriteLine("OnAllSpreadsheetsLoaded begins");
        AllSpreadsheetsLoaded?.Invoke(this, EventArgs.Empty);
    }

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

    internal virtual SpreadCellEditEndedEventArgs RaiseCellEditEnded(Spreadsheet spreadsheet, SpreadCellEditEndedEventArgs eventArgs)
    {
        CellEditEnded?.Invoke(this, eventArgs);

        if (eventArgs.Cancel) return eventArgs;

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

    #region Initialization.

    public Spreadbook()
    {
        InitializeComponent();

        bookTabControl = this.FindControl<TabControl>(nameof(SpreadsheetsTabControl));
        spreadsheets = new List<Spreadsheet>();
    }

    #endregion

    #region Spreadsheets handling.

    public Spreadsheet FindSpreadsheet(string sheetName)
    {
        return spreadsheets.FirstOrDefault(s => s.WorksheetName == sheetName);
    }

    public Spreadsheet FindSpreadsheet(int sheetIndex)
    {
        return spreadsheets.FirstOrDefault(s => s.WorksheetIndex == sheetIndex);
    }

    protected TabItem FindTabItem(Spreadsheet spreadsheet)
    {
        var tab =
            bookTabControl
                .Items
                .FirstOrDefault(t => (t as TabItem)?.Content as Spreadsheet == spreadsheet)
                as TabItem;

        return tab;
    }

    public void SelectSpreadsheet(string sheetName)
    {
        var spreadsheet = 
            FindSpreadsheet(sheetName) ?? 
            throw new ArgumentException($"Worksheet '{sheetName}' not found.");

        var tab = 
            FindTabItem(spreadsheet) ?? 
            throw new ArgumentException($"Worksheet '{sheetName}' not found.");

        tab.IsSelected = true;
    }

    public void SelectSpreadsheet(int sheetIndex)
    {
        var spreadsheet = 
            FindSpreadsheet(sheetIndex) ?? 
            throw new ArgumentException($"Worksheet #{sheetIndex} not found.");

        var tab = 
            FindTabItem(spreadsheet) ?? 
            throw new ArgumentException($"Worksheet #{sheetIndex} not found.");

        tab.IsSelected = true;
    }

    #endregion
}
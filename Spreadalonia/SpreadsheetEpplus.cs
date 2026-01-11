using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;
using OfficeOpenXml;

namespace Spreadalonia;

public class SpreadsheetEpplus : Spreadsheet
{
    public ExcelWorksheet EpplusWorksheet { get; set; }

    public SpreadbookEpplus SpreadBookEpplus => SpreadBook as SpreadbookEpplus;

    public override int WorksheetIndex
    {
        get => EpplusWorksheet.Index;
        set => throw new InvalidOperationException("Cannot change worksheet index directly.");
    }

    public override string WorksheetName
    {
        get => EpplusWorksheet.Name;
        set => throw new InvalidOperationException("Cannot change worksheet name directly.");
    }

    public object GetCellValue(int column, int row)
    {
        return EpplusWorksheet.Cells[row + 1, column + 1].Value;
    }

    private ExcelRange GetEpplusCell(SpreadCell cell)
    {
        return EpplusWorksheet.Cells[cell.Row + 1, cell.Column + 1];
    }

    internal List<KeyValuePair<(int, int), string>> GetDataFromEpplus()
    {
        var dataForSpreadsheet = EpplusWorksheet.Cells
            .Where(c => c.Value != null && !c.EntireRow.Hidden && !c.EntireColumn.Hidden)
            .Select(EpplusCellToSpreadsheetCell)
            .ToList();

        return dataForSpreadsheet;
    }

    private static KeyValuePair<(int, int), string> EpplusCellToSpreadsheetCell(ExcelRangeBase c)
    {
        var spreadsheetCells = new KeyValuePair<(int, int), string>(
            key: new ValueTuple<int, int>(c.EntireColumn.StartColumn - 1, c.EntireRow.StartRow - 1),
            value: c.Text);

        return spreadsheetCells;
    }

    internal void SetEpplusCellValue(SpreadCell cell)
    {
        SetEpplusCellValue(cell, cell.TextValue);
    }

    internal void SetEpplusCellValue(SpreadCell cell, string newText)
    {
        var epplusCell = GetEpplusCell(cell);

        // TODO: how to get cell data type?
        if (double.TryParse(newText, out double parsedNumber))
        {
            epplusCell.Value = parsedNumber;
            return;
        }

        if (DateTime.TryParse(newText, out DateTime parsedDate))
        {
            epplusCell.Value = parsedDate;
            return;
        }

        epplusCell.Value = newText;
    }

    /// <summary>
    /// Set cell text from specified value based on cell format.
    /// </summary>
    /// <param name="column">0-based cell column index.</param>
    /// <param name="row">0-based cell row index.</param>
    /// <param name="value">Desired value. May be string, DateTime or double.</param>
    public void SetCellValue(int column, int row, object value)
    {
        var epplusCell = EpplusWorksheet.Cells[row + 1, column + 1];
        epplusCell.Value = value;
        SkipHandlers = true;
        SetCellText(column, row, epplusCell.Text);
        SkipHandlers = false;
    }

    public void CopyRow(int sourceRowIndex, int destinationRowIndex, ExcelRangeCopyOptionFlags flags)
    {
        base.CopyRow(sourceRowIndex, destinationRowIndex);

        var sourceRowEpplus = sourceRowIndex + 1;
        var destinationRowEpplus = destinationRowIndex + 1;
        var columnsCount = EpplusWorksheet.Dimension.Columns;

        var sourceRowCells = 
            EpplusWorksheet
            .Cells[sourceRowEpplus, 1, sourceRowEpplus, columnsCount];

        var destinationRowCells = 
            EpplusWorksheet
            .Cells[destinationRowEpplus, 1, destinationRowEpplus, columnsCount];

        sourceRowCells.Copy(destinationRowCells, flags);

        EpplusWorksheet.Row(destinationRowEpplus).StyleID =
            EpplusWorksheet.Row(sourceRowEpplus).StyleID;

        SpreadBookEpplus.ReloadDataFromEpplus();
    }

    public override void CopyRow(int sourceRowIndex, int destinationRowIndex)
    {
        CopyRow(sourceRowIndex, destinationRowIndex, 0);
    }

    public override void InsertRows(SelectionRange selectedRows)
    {
        base.InsertRows(selectedRows);

        // Index of first new row.
        var firstInsertedRowIndex = selectedRows.Bottom;
        var firstInsertedRowEpplus = firstInsertedRowIndex + 1;
        var insertedRowsCount = selectedRows.Height;

        EpplusWorksheet.InsertRow(
            rowFrom: firstInsertedRowEpplus,
            rows: insertedRowsCount,
            copyStylesFromRow: firstInsertedRowEpplus);

        SpreadBookEpplus.ReloadDataFromEpplus();
    }

    public void InsertRowWithFormulas()
    {
        if (Selection.Count != 1) return;

        InsertRowWithFormulas(Selection[0].Top);
    }

    public void InsertRowWithFormulas(int insertBeforeThisRowIndex)
    {
        var selection = new SelectionRange(0, insertBeforeThisRowIndex);
        base.InsertRows(selection);

        EpplusHelper.InsertRowWithFormulas(EpplusWorksheet, insertBeforeThisRowIndex + 1);

        SpreadBookEpplus.ReloadDataFromEpplus();
    }

    public void AppendRowWithFormulas()
    {
        if (Selection.Count != 1) return;

        AppendRowWithFormulas(Selection[0].Top);
    }

    public void AppendRowWithFormulas(int appendAfterThisRowIndex)
    {
        var selection = new SelectionRange(0, appendAfterThisRowIndex + 1);
        base.InsertRows(selection);

        EpplusHelper.AppendRowWithFormulas(EpplusWorksheet, appendAfterThisRowIndex + 1);

        SpreadBookEpplus.ReloadDataFromEpplus();
    }

    public override void DeleteRows(SelectionRange selection)
    {
        base.DeleteRows(selection);

        EpplusWorksheet.DeleteRow(selection.Top + 1, selection.Height);

        SpreadBookEpplus.ReloadDataFromEpplus();
    }

    protected override void OnClearContentsEnded(List<SpreadCell> cellsToClear)
    {
        base.OnClearContentsEnded(cellsToClear);

        foreach (var spreadCell in cellsToClear)
        {
            GetEpplusCell(spreadCell).Clear();
        }

        SpreadBookEpplus.ReloadDataFromEpplus();
    }

    protected override void OnCellEditEnded(SpreadCell cell)
    {
        base.OnCellEditEnded(cell);

        SetEpplusCellValue(cell);

        SpreadBookEpplus.ReloadDataFromEpplus();
    }

    protected override void OnSetDataEnded(List<SpreadCell> newData)
    {
        base.OnSetDataEnded(newData);

        foreach (var cell in newData)
        {
            SetEpplusCellValue(cell);
        }

        SpreadBookEpplus.ReloadDataFromEpplus();
    }
}
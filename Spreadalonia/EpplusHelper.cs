using System.Linq;
using OfficeOpenXml;

namespace Spreadalonia;

internal class EpplusHelper
{
    public static void InsertRowWithFormulas(ExcelWorksheet epplusWorksheet, int rowIndexEpplus)
    {
        var newRowIndexEpplus = rowIndexEpplus;
        var sourceRowIndexEpplus = rowIndexEpplus + 1;

        epplusWorksheet.InsertRow(
            rowFrom: newRowIndexEpplus,
            rows: 1,
            copyStylesFromRow: sourceRowIndexEpplus);

        CopyRowFormulas(epplusWorksheet, sourceRowIndexEpplus, newRowIndexEpplus);
    }

    public static void AppendRowWithFormulas(ExcelWorksheet epplusWorksheet, int rowIndexEpplus)
    {
        var newRowIndexEpplus = rowIndexEpplus + 1;
        var sourceRowIndexEpplus = rowIndexEpplus;

        epplusWorksheet.InsertRow(
            rowFrom: newRowIndexEpplus,
            rows: 1,
            copyStylesFromRow: sourceRowIndexEpplus);

        CopyRowFormulas(epplusWorksheet, sourceRowIndexEpplus, newRowIndexEpplus);
    }

    private static void CopyRowFormulas(ExcelWorksheet epplusWorksheet, int sourceRowIndexEpplus, int destinationRowIndexEpplus)
    {
        var columnsCount = epplusWorksheet.Dimension.Columns;

        // Copy only formulas cell by cell.
        var sourceCellsWithFormula =
            epplusWorksheet
                .Cells[sourceRowIndexEpplus, 1, sourceRowIndexEpplus, columnsCount]
                .Where(c => !string.IsNullOrEmpty(c.Formula));

        foreach (var sourceCell in sourceCellsWithFormula)
        {
            var destinationCell =
                epplusWorksheet.Cells[
                    destinationRowIndexEpplus,
                    sourceCell.EntireColumn.StartColumn];

            sourceCell.Copy(destinationCell);
        }
    }
}
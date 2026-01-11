using System;
using System.Collections.Generic;
using Avalonia.Media;

namespace Spreadalonia;

/// <summary>
/// <see cref="EventArgs"/> for the <see cref="Spreadsheet.CellSizeChanged"/> event.
/// </summary>
public class CellSizeChangedEventArgs : EventArgs
{
    /// <summary>
    /// The new width of the cell.
    /// </summary>
    public double Width { get; }

    /// <summary>
    /// The new height of the cell.
    /// </summary>
    public double Height { get; }

    /// <summary>
    /// The horizontal coordinate of the cell.
    /// </summary>
    public int Left { get; }

    /// <summary>
    /// The vertical coordinate of the cell.
    /// </summary>
    public int Top { get; }

    internal CellSizeChangedEventArgs(int left, int top, double width, double height) : base()
    {
        this.Left = left;
        this.Top = top;
        this.Width = width;
        this.Height = height;
    }
}

/// <summary>
/// <see cref="EventArgs"/> for the <see cref="Spreadsheet.ColorDoubleTapped"/> event.
/// </summary>
public class ColorDoubleTappedEventArgs : EventArgs
{
    /// <summary>
    /// The horizontal coordinate of the cell.
    /// </summary>
    public int Left { get; }

    /// <summary>
    /// The vertical coordinate of the cell.
    /// </summary>
    public int Top { get; }

    /// <summary>
    /// The colour contained in the cell.
    /// </summary>
    public Color Color { get; }

    /// <summary>
    /// Set this to <see langword="true"/> to signal that the event has been completely handled.
    /// </summary>
    public bool Handled { get; set; }

    internal ColorDoubleTappedEventArgs(int left, int top, Color color) : base()
    {
        Left = left;
        Top = top;
        Color = color;
        Handled = false;
    }
}

public class SpreadCellEditStartedEventArgs : EventArgs
{
    public SpreadCell Cell { get; }

    /// <summary>
    /// Should be editing canceled (true) or continued (false)?
    /// </summary>
    public bool Cancel { get; set; }

    public SpreadCellEditStartedEventArgs(SpreadCell cell)
    {
        Cell = cell;
    }
}

public class SpreadCellEditEndedEventArgs : EventArgs
{
    public SpreadCell Cell { get; }

    /// <summary>
    /// New text for cell.
    /// </summary>
    public string NewText { get; }

    /// <summary>
    /// Should be new text canceled (true) or accepted (false)?
    /// </summary>
    public bool Cancel { get; set; }

    public SpreadCellEditEndedEventArgs(SpreadCell cell, string newText)
    {
        Cell = cell;
        NewText = newText;
    }
}

public class MassiveSpreadCellsEditStartedEventArgs : EventArgs
{
    public List<SpreadCell> OldCellsContent { get; }

    public List<SpreadCell> NewCellsContent { get; }

    public IEnumerable<SelectionRange> Range { get; }

    /// <summary>
    /// Should be editing canceled (true) or continued (false)?
    /// </summary>
    public bool Cancel { get; set; }

    public MassiveSpreadCellsEditStartedEventArgs(List<SpreadCell> oldCellsContent, List<SpreadCell> newCellsContent, IEnumerable<SelectionRange> range)
    {
        OldCellsContent = oldCellsContent;
        NewCellsContent = newCellsContent;
        Range = range;
    }
}

public class RowCopiedEventArgs : EventArgs
{
    public int SourceRowIndex { get; set; }

    public int DestinationRowIndex { get; set; }

    public RowCopiedEventArgs(int sourceRowIndex, int destinationRowIndex)
    {
        SourceRowIndex = sourceRowIndex;
        DestinationRowIndex = destinationRowIndex;
    }
}

public class RowsAddedEventArgs : EventArgs
{
    public int FirstAddedRowIndex { get; set; }

    public int AddedRowsCount { get; set; }

    public RowsAddedEventArgs(int firstAddedRowIndex, int addedRowsCount)
    {
        FirstAddedRowIndex = firstAddedRowIndex;
        AddedRowsCount = addedRowsCount;
    }
}

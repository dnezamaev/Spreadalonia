namespace Spreadalonia;

/// <summary>
/// Configurations of Spreadsheet.
/// Is*Enabled options affect only user actions
/// like context menu and keyboard input,
/// public methods of Spreadsheet class still available.
/// </summary>
public class SpreadsheetOptions
{
    /// <summary>
    /// Allow undo actions?
    /// </summary>
    public bool IsUndoEnabled { get; set; } = true;

    /// <summary>
    /// Allow redo actions?
    /// </summary>
    public bool IsRedoEnabled { get; set; } = true;

    /// <summary>
    /// Allow copy cell content?
    /// </summary>
    public bool IsCopyEnabled { get; set; } = true;

    /// <summary>
    /// Allow cut cell content?
    /// </summary>
    public bool IsCutEnabled { get; set; } = true;

    /// <summary>
    /// Allow paste cell content?
    /// </summary>
    public bool IsPasteEnabled { get; set; } = true;

    /// <summary>
    /// Allow clear cell content?
    /// </summary>
    public bool IsClearContentsEnabled { get; set; } = true;

    /// <summary>
    /// Allow insert columns?
    /// </summary>
    public bool IsInsertColumnsEnabled { get; set; } = true;

    /// <summary>
    /// Allow insert rows?
    /// </summary>
    public bool IsInsertRowsEnabled { get; set; } = true;

    /// <summary>
    /// Allow remove columns?
    /// </summary>
    public bool IsDeleteColumnsEnabled { get; set; } = true;

    /// <summary>
    /// Allow remove rows?
    /// </summary>
    public bool IsDeleteRowsEnabled { get; set; } = true;

    /// <summary>
    /// Allow reset cell format?
    /// </summary>
    public bool IsResetFormatEnabled { get; set; } = true;
}
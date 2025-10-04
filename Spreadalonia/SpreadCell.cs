using System.Drawing;

namespace Spreadalonia
{
    public class SpreadCell
    {
        public Spreadsheet Worksheet { get; set; }

        public int Column { get; set; }

        public int Row { get; set; }

        public string TextValue { get; set; }

        public SpreadCell() { }

        public SpreadCell(Spreadsheet worksheet, int column, int row, string textValue = null)
        {
            Column = column;
            Row = row;
            Worksheet = worksheet;
            TextValue = textValue;
        }

        public SpreadCell(Spreadsheet worksheet, (int column, int row) addressTuple, string textValue = null)
            : this(worksheet, addressTuple.column, addressTuple.row, textValue) { }
    }
}

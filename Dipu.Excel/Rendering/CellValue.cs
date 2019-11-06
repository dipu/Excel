namespace Dipu.Excel.Rendering
{
    public class CellValue
    {
        public int Row;

        public int Column;

        public object Value;

        public CellValue(int row, int column, object value = null)
        {
            this.Row = row;
            this.Column = column;
            this.Value = value;
        }

        public override string ToString()
        {
            return $"[{Row},{Column}] = {Value}";
        }
    }
}
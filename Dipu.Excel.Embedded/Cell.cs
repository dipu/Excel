namespace Dipu.Excel.Embedded
{
    public class Cell : ICell
    {
        private object value;
        private Style style;

        public Cell(Worksheet worksheet, int row, int column)
        {
            Worksheet = worksheet;
            Row = row;
            Column = column;
        }

        IWorksheet ICell.Worksheet => this.Worksheet;

        public Worksheet Worksheet { get; }

        int ICell.Row => this.Row;

        public int Row { get; }

        int ICell.Column => this.Column;

        public int Column { get; }

        object ICell.Value { get => this.Value; set => this.Value = value; }

        public object Value
        {
            get => this.value;
            set
            {
                if (this.UpdateValue(value))
                {
                    this.Worksheet.AddDirtyValue(this);
                }
            }
        }

        public Style Style
        {
            get => style;
            set
            {
                if (!Equals(this.style, value))
                {
                    this.style = value;
                    this.Worksheet.AddDirtyStyle(this);
                }
            }
        }
        
        public bool UpdateValue(object newValue)
        {
            if (!Equals(this.value, newValue))
            {
                this.value = newValue;
                return true;
            }

            return false;
        }
    }
}

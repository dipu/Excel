using System;

namespace Dipu.Excel.Headless
{
    public class Cell : ICell
    {
        public Cell(Worksheet worksheet, Row row, Column column)
        {
            this.Worksheet = worksheet;
            this.Row = row;
            this.Column = column;
        }

        public IWorksheet Worksheet { get; }

        IRow ICell.Row => this.Row;

        public Row Row { get; }

        IColumn ICell.Column => this.Column;

        public Column Column { get; }

        public object Value { get; set; }

        public string Comment { get; set; }

        public Style Style { get; set; }

        public string NumberFormat { get; set; }

        public IValueConverter ValueConverter { get; set; }

        public Range Options { get; set; }

        public bool IsRequired { get; set; }

        public void Clear()
        {
            this.Value = string.Empty;
            this.NumberFormat = null;
            this.Style = null;
        }
    }
}

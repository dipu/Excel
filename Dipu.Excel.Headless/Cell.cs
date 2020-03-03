using System;

namespace Dipu.Excel.Headless
{
    public class Cell : ICell
    {
        public Cell(Worksheet worksheet, int row, int column)
        {
            Worksheet = worksheet;
            Row = row;
            Column = column;
        }

        public IWorksheet Worksheet { get; }

        public int Row { get; }
        
        public int Column { get; }
        
        public object Value { get; set; }
        
        public string Comment { get; set; }
        
        public Style Style { get; set; }
        
        public string NumberFormat { get; set; }
        
        public IValueConverter ValueConverter { get; set; }

        public Range Options { get; set; }

        public void Clear()
        {
            this.Value = string.Empty;
            this.NumberFormat = null;
            this.Style = null;
        }
    }
}

using System;

namespace Dipu.Excel.Embedded
{
    public class Cell : ICell
    {
        private object value;
        private Style style;
        private string numberFormat;

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

        public string NumberFormat
        {
            get => numberFormat;
            set
            {
                if (!Equals(this.numberFormat, value))
                {
                    this.numberFormat = value;
                    this.Worksheet.AddDirtyNumberFormat(this);
                }
            }
        }

        public bool UpdateValue(object newValue)
        {
            bool update;

            if (this.value is decimal @decimal && newValue is double @double)
            {
                const double decimalMin = (double)decimal.MinValue;
                const double decimalMax = (double)decimal.MaxValue;

                if (@double < decimalMin || @double > decimalMax)
                {
                    update = true;
                }
                else
                {
                    update = ((decimal)@double).CompareTo(@decimal) != 0;
                }
            }
            else
            {
                update = !Equals(this.value, newValue);
            }
            
            if (update)
            {
                this.value = newValue;
            }

            return update;
        }
    }
}

namespace Dipu.Excel.Embedded
{
    public class Cell : ICell
    {
        private object value;
        private Style style;
        private string numberFormat;
        private IExcelValueConverter excelValueConverter;
        private readonly IExcelValueConverter defaultExcelValueConverter = new DefaultExcelConverter();
        private string comment;

        public Cell(IEmbeddedWorksheet worksheet, int row, int column)
        {
            Worksheet = worksheet;
            Row = row;
            Column = column;
        }

        IWorksheet ICell.Worksheet => this.Worksheet;

        public IEmbeddedWorksheet Worksheet { get; }

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
                if (!Equals(this.value, value))
                {
                    this.Worksheet.AddDirtyValue(this);
                    this.value = value;
                }
            }
        }

        public string Comment
        {
            get => comment;
            set
            {
                if (!Equals(this.comment, value))
                {
                    this.comment = value;
                    this.Worksheet.AddDirtyComment(this);
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

        public IExcelValueConverter ExcelValueConverter
        {
            get => excelValueConverter ?? this.defaultExcelValueConverter;
            set => excelValueConverter = value;
        }

        public override string ToString()
        {
            return $"{Row}:{Column}";
        }

        public bool UpdateValue(object rawExcelValue)
        {
            var excelValue = this.ExcelValueConverter.Convert(this, rawExcelValue);
            var update = !Equals(this.value, excelValue);

            if (update)
            {
                this.value = excelValue;
            }

            return update;
        }
    }
}

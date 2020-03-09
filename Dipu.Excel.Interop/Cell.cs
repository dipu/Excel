namespace Dipu.Excel.Embedded
{
    public class Cell : ICell
    {
        private object value;
        private Style style;
        private string numberFormat;
        private Range options;

        private IValueConverter valueConverter;
        private readonly IValueConverter defaultValueConverter = new DefaultValueConverter();
        private string comment;

        public Cell(IEmbeddedWorksheet worksheet, Row row, Column column)
        {
            Worksheet = worksheet;
            Row = row;
            Column = column;
        }

        IWorksheet ICell.Worksheet => this.Worksheet;

        public IEmbeddedWorksheet Worksheet { get; }

        IRow ICell.Row => this.Row;

        public Row Row { get; }

        IColumn ICell.Column => this.Column;

        public Column Column { get; }

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
                if (!this.style?.Equals(value) ?? value != null)
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

        public IValueConverter ValueConverter
        {
            get => valueConverter ?? this.defaultValueConverter;
            set => valueConverter = value;
        }

        public Range Options
        {
            get => options;
            set
            {
                if (!Equals(this.options, value))
                {
                    this.options = value;
                    this.Worksheet.AddDirtyOptions(this);
                }
            }
        }

        public override string ToString()
        {
            return $"{Row}:{Column}";
        }

        public bool UpdateValue(object rawExcelValue)
        {
            var excelValue = this.ValueConverter.Convert(this, rawExcelValue);
            var update = !Equals(this.value, excelValue);

            if (update)
            {
                this.value = excelValue;
            }

            return update;
        }

        public void Clear()
        {
            this.Value = string.Empty;
            this.Style = null;
            this.NumberFormat = null;
        }
    }
}

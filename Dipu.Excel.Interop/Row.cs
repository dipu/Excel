using System;

namespace Dipu.Excel.Embedded
{
    public class Row : IRow, IComparable<Row>
    {
        private bool hidden;

        public Row(Worksheet worksheet, int index)
        {
            Worksheet = worksheet;
            Index = index;
        }

        IWorksheet IRow.Worksheet => this.Worksheet;

        public Worksheet Worksheet { get; }

        public int Index { get; }

        bool IRow.Hidden { get => this.Hidden; set => this.Hidden = value; }

        public bool Hidden
        {
            get => hidden;
            set
            {
                hidden = value;
                this.Worksheet.AddDirtyRow(this);
            }
        }

        int IRow.Index { get; }

        public int CompareTo(Row other)
        {
            return this.Index.CompareTo(other.Index);
        }
    }
}

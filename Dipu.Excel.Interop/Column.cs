using System;

namespace Dipu.Excel.Embedded
{
    public class Column : IColumn, IComparable<Column>
    {
        public Column(Worksheet worksheet, int index)
        {
            Worksheet = worksheet;
            Index = index;
        }

        public IWorksheet Worksheet { get; }

        public int Index { get; }

        public int CompareTo(Column other)
        {
            return this.Index.CompareTo(other.Index);
        }
    }
}

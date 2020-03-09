using System;

namespace Dipu.Excel.Embedded
{
    public class Row : IRow, IComparable<Row>
    {
        public Row(Worksheet worksheet, int index)
        {
            Worksheet = worksheet;
            Index = index;
        }

        public IWorksheet Worksheet { get; }

        public int Index { get; }

        public int CompareTo(Row other)
        {
            return this.Index.CompareTo(other.Index); 
        }
    }
}

namespace Dipu.Excel.Headless
{
    public class Column : IColumn
    {
        public Column(Worksheet worksheet, int index)
        {
            Worksheet = worksheet;
            Index = index;
        }

        public IWorksheet Worksheet { get; }

        public int Index { get; }
    }
}

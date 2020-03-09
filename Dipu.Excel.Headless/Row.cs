namespace Dipu.Excel.Headless
{
    public class Row : IRow
    {
        public Row(Worksheet worksheet, int index)
        {
            Worksheet = worksheet;
            Index = index;
        }

        public IWorksheet Worksheet { get; }

        public int Index { get; }

        public bool Hidden { get; set; }
    }
}

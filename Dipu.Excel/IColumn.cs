namespace Dipu.Excel
{
    public interface IColumn
    {
        IWorksheet Worksheet { get; }

        int Index { get; }
    }
}

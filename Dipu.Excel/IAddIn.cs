namespace Dipu.Excel
{
    public interface IAddIn
    {
        IWorkbook[] Workbooks { get; }
    }
}

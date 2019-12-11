namespace Dipu.Excel
{
    public interface IWorkbook
    {
        bool IsActive { get; }

        IWorksheet[] Worksheets { get; }

        void Close(bool? saveChanges = null, string fileName = null);

        IWorksheet AddWorksheet(int? index = null, IWorksheet before = null, IWorksheet after = null);
    }
}

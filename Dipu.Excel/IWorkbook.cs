using System.Collections;

namespace Dipu.Excel
{
    public interface IWorkbook
    {
        IWorksheet CreateSheet(int? index = null, IWorksheet before = null, IWorksheet after = null);

        IWorksheet[] Worksheets { get; }

        void Close(bool? saveChanges = null, string fileName = null);
    }
}

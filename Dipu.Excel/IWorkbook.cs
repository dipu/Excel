using System.Collections;
using System.Threading.Tasks;

namespace Dipu.Excel
{
    public interface IWorkbook
    {
        IWorksheet CreateSheet(int? index = null, IWorksheet before = null, IWorksheet after = null);

        IWorksheet[] Worksheets { get; }

        bool Active { get; }

        void Close(bool? saveChanges = null, string fileName = null);
    }
}

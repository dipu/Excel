using System;
using System.Threading.Tasks;

namespace Dipu.Excel
{
    public interface IWorksheet
    {
        event EventHandler<CellChangedEvent> CellsChanged;

        IWorkbook Workbook { get; }

        string Name { get; set; }

        bool IsActive { get; }

        ICell this[int row, int column]
        {
            get;
        }

        Task Flush();
    }
}

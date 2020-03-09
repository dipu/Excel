using System;
using System.Threading.Tasks;

namespace Dipu.Excel
{
    public interface IWorksheet
    {
        event EventHandler<CellChangedEvent> CellsChanged;

        event EventHandler<string> SheetActivated;

        IWorkbook Workbook { get; }

        string Name { get; set; }

        bool IsActive { get; }

        IRow Row(int index);

        IColumn Column(int index);

        ICell this[int row, int column]
        {
            get;
        }

        Task Flush();
    }
}

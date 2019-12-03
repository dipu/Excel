using System;
using System.Threading.Tasks;

namespace Dipu.Excel
{
    public interface IWorksheet
    {
        event EventHandler<CellChangedEvent> CellsChanged;

        string Name { get; set; }

        ICell this[int row, int column]
        {
            get;
        }

        Task Flush();
    }
}

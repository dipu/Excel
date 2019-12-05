using System;
using System.Threading.Tasks;

namespace Dipu.Excel
{
    public interface IWorksheet
    {
        event EventHandler<CellChangedEvent> CellsChanged;

        string Name { get; set; }

        bool Active { get; }

        ICell this[int row, int column]
        {
            get;
        }

        Task Flush();
    }
}

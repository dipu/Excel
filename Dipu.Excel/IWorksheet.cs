using System;
using System.Threading.Tasks;

namespace Dipu.Excel
{
    public interface IWorksheet
    {
        event EventHandler<CellChangedEvent> CellChanged;

        string Name { get; set; }

        ICell this[int row, int column]
        {
            get;
        }

        Task Flush();
    }
}

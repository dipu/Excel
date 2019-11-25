using System;

namespace Dipu.Excel
{
    public class CellChangedEvent : EventArgs
    {
        public ICell[] Cells { get; }

        public CellChangedEvent(ICell[] cells)
        {
            Cells = cells;
        }
    }
}
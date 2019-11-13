using System;

namespace Dipu.Excel
{
    public class CellChangedEvent : EventArgs
    {
        public CellChangedEvent(int row, int column)
        {
            Row = row;
            Column = column;
        }

        public int Row { get; }

        public int Column { get; }
    }
}
using System;

namespace Dipu.Excel
{
    public class Binding : IBinding
    {
        private readonly Action<ICell> toCell;

        private readonly Action<ICell> toDomain;

        public bool OneWayBinding => toDomain == null;

        public bool TwoWayBinding => !this.OneWayBinding;

        public object Value { get; }

        public Binding(Action<ICell> toCell = null, Action<ICell> toDomain = null)
        {
            this.toCell = toCell;
            this.toDomain = toDomain;
        }

        public void ToCell(ICell cell)
        {
            this?.toCell(cell);
        }

        public void ToDomain(ICell cell)
        {
            this?.toDomain(cell);
        }
    }
}
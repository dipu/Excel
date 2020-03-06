namespace Dipu.Excel
{
    using System;

    public class Binding : IBinding
    {
        private readonly Action<ICell> toCell;

        private readonly Action<ICell> toDomain;

        public Binding(Action<ICell> toCell = null, Action<ICell> toDomain = null)
        {
            this.toCell = toCell;
            this.toDomain = toDomain;
        }

        public bool OneWayBinding => this.toDomain == null;

        public bool TwoWayBinding => !this.OneWayBinding;

        public object Value { get; }

        public void ToCell(ICell cell)
        {
            this.toCell?.Invoke(cell);
        }

        public void ToDomain(ICell cell)
        {
            this.toDomain?.Invoke(cell);
        }
    }
}
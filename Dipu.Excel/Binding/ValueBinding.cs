namespace Dipu.Excel
{
    public class ValueBinding : IBinding
    {
        public bool OneWayBinding => true;

        public bool TwoWayBinding => !this.OneWayBinding;

        public object Value { get; }

        public ValueBinding(object value)
        {
            this.Value = value;
        }

        public void ToCell(ICell cell)
        {
            cell.Value = this.Value;
        }

        public void ToDomain(ICell cell)
        {
        }
    }
}
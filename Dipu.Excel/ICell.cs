namespace Dipu.Excel
{
    public interface ICell
    {
        IWorksheet Worksheet { get; }

        int Row { get; }

        int Column { get; }
                
        object Value { get; set; }

        Range Options { get; set; }

        string Comment { get; set; }

        Style Style { get; set; }

        string NumberFormat { get; set; }

        IValueConverter ValueConverter { get; set; }

        void Clear();
    }
}

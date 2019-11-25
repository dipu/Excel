namespace Dipu.Excel
{
    public interface ICell
    {
        IWorksheet Worksheet { get; }

        int Row { get; }

        int Column { get; }

        object Value { get; set; }

        Style Style { get; set; }

        string NumberFormat { get; set; }
    }
}

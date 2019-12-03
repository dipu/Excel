namespace Dipu.Excel
{
    public interface IValueConverter
    {
        object Convert(ICell cell, object excelValue);
    }
}

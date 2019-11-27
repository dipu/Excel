namespace Dipu.Excel
{
    public interface IExcelValueConverter
    {
        object Convert(ICell cell, object excelValue);
    }
}

using System;
using System.Linq;

namespace Dipu.Excel
{
    public class DefaultExcelConverter : IExcelValueConverter
    {
        public object Convert(ICell cell, object excelValue)
        {
            {
                if (cell.Value is decimal @decimal && excelValue is double @double)
                {
                    const double min = (double)decimal.MinValue;
                    const double max = (double)decimal.MaxValue;

                    if (@double < min)
                    {
                        return min;
                    }

                    if (@double > max)
                    {
                        return max;
                    }

                    return System.Convert.ToDecimal(excelValue);
                }
            }

            {
                if (cell.Value is int @integer && excelValue is double @double)
                {
                    const double min = (double)int.MinValue;
                    const double max = (double)int.MaxValue;

                    if (@double < min)
                    {
                        return min;
                    }

                    if (@double > max)
                    {
                        return max;
                    }

                    return System.Convert.ToInt32(excelValue);
                }
            }

            {
                if (cell.Value is DateTime dateTime && excelValue is double @double)
                {
                    return DateTime.FromOADate(@double);
                }
            }

            {
                if (cell.Value is string @string)
                {
                    if (excelValue == null)
                    {
                        return string.Empty;
                    }

                    if (!(excelValue is string))
                    {
                        return excelValue.ToString();
                    }
                }
            }

            return excelValue;
        }
    }
}

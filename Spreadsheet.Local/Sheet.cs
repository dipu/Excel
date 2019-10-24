using Microsoft.Office.Interop.Excel;

namespace Dipu.Spreadsheet.Local
{
    public class Sheet : ISheet
    {
        public Sheet(Workbook workbook, Worksheet comWorksheet)
        {
            this.Workbook = workbook;
            this.ComWorksheet = comWorksheet;
        }

        public Workbook Workbook { get; set; }

        public Worksheet ComWorksheet { get; set; }

        public string Name { get => this.ComWorksheet.Name; set => this.ComWorksheet.Name = value; }
    }
}

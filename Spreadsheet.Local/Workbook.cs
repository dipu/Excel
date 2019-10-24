using System.Collections.Generic;

namespace Dipu.Spreadsheet.Local
{
    public class Workbook : IWorkbook
    {

        public Workbook(Host host, Microsoft.Office.Interop.Excel.Workbook comWorkbook)
        {
            this.Host = host;
            this.ComWorkbook = comWorkbook;
        }

        public Host Host { get; }

        public Microsoft.Office.Interop.Excel.Workbook ComWorkbook { get; }

        public IEnumerable<ISheet> Sheets { get; }

        public ISheet CreateSheet()
        {
            var index = this.ComWorkbook.Sheets.Add();
            var comWorksheet = this.ComWorkbook.Sheets[index];
            return new Sheet(this, comWorksheet);
        }
    }
}

using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace Dipu.Excel.Embedded
{
    public class Worksheet : IWorksheet
    {
        public Worksheet(Workbook workbook, ComWrapper<InteropWorksheet> interopWorksheet)
        {
            this.Workbook = workbook;
            this.InteropWorksheet = interopWorksheet;
        }
        
        public Workbook Workbook { get; set; }

        public ComWrapper<InteropWorksheet> InteropWorksheet { get; set; }

        public string Name { get => this.InteropWorksheet.ComObject.Name; set => this.InteropWorksheet.ComObject.Name = value; }

        public int Index => this.InteropWorksheet.ComObject.Index;
    }
}

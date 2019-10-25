using System.Collections.Generic;
using System.Linq;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace Dipu.Excel.Embedded
{
    public class Workbook : IWorkbook
    {
        private readonly Dictionary<ComWrapper<InteropWorksheet>, Worksheet> worksheetByInteropWorksheet;

        public Workbook(AddIn addIn, ComWrapper<InteropWorkbook> interopWorkbook)
        {
            this.AddIn = addIn;
            this.InteropWorkbook = interopWorkbook;
            this.worksheetByInteropWorksheet = new Dictionary<ComWrapper<InteropWorksheet>, Worksheet>();
        }

        public AddIn AddIn { get; }

        public ComWrapper<InteropWorkbook> InteropWorkbook { get; }

        public IWorksheet CreateSheet(int? index, IWorksheet before = null, IWorksheet after = null)
        {
            ComWrapper<InteropWorksheet> interopWorksheet;

            if (index.HasValue && index.Value == 0)
            {
                interopWorksheet = Com.Wrap((InteropWorksheet)this.InteropWorkbook.ComObject.Sheets.Add());
            }
            else
            {
                if (before != null)
                {
                    interopWorksheet = Com.Wrap((InteropWorksheet)this.InteropWorkbook.ComObject.Sheets.Add(((Worksheet)before).InteropWorksheet.ComObject));
                }
                else if (after != null)
                {
                    interopWorksheet = Com.Wrap((InteropWorksheet)this.InteropWorkbook.ComObject.Sheets.Add(null, ((Worksheet)after).InteropWorksheet.ComObject));
                }
                else
                {
                    var sortedWorksheets = this.worksheetByInteropWorksheet.OrderBy(v => v.Value.Index).Select(v => v.Key).ToArray();
                    ComWrapper<InteropWorksheet> append = null;
                    if (sortedWorksheets.Any())
                    {
                        if (!index.HasValue || index > sortedWorksheets.Length - 1)
                        {
                            index = sortedWorksheets.Length - 1;
                        }

                        append = sortedWorksheets[index.Value];
                    }

                    interopWorksheet = Com.Wrap((InteropWorksheet)this.InteropWorkbook.ComObject.Sheets.Add(null, append?.ComObject));
                }
            }
            
            var worksheet = new Worksheet(this, interopWorksheet);
            this.worksheetByInteropWorksheet.Add(interopWorksheet, worksheet);
            return worksheet;
        }

        public IWorksheet[] Worksheets => this.worksheetByInteropWorksheet.Values.Cast<IWorksheet>().ToArray();

        public Worksheet New(ComWrapper<InteropWorksheet> interopWorksheet)
        {
            var worksheet = new Worksheet(this, interopWorksheet);
            this.worksheetByInteropWorksheet.Add(interopWorksheet, worksheet);
            return worksheet;
        }
    }
}

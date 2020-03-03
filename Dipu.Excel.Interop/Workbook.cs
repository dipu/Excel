namespace Dipu.Excel.Embedded
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
    using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

    public class Workbook : IWorkbook
    {
        private readonly Dictionary<InteropWorksheet, Worksheet> worksheetByInteropWorksheet;

        public Workbook(AddIn addIn, InteropWorkbook interopWorkbook)
        {
            this.AddIn = addIn;
            this.InteropWorkbook = interopWorkbook;
            this.worksheetByInteropWorksheet = new Dictionary<InteropWorksheet, Worksheet>();
            this.AddIn.Application.WorkbookNewSheet += this.ApplicationOnWorkbookNewSheet;
            this.AddIn.Application.SheetBeforeDelete += this.ApplicationOnSheetBeforeDelete;
        }

        public AddIn AddIn { get; }

        public InteropWorkbook InteropWorkbook { get; }

        /// <summary>
        /// Return a Zero-Based Row, Column NamedRanges
        /// </summary>
        /// <returns></returns>
        public List<NamedRange> GetNamedRanges()
        {
            var ranges = new List<NamedRange>();

            Microsoft.Office.Interop.Excel.Names names = this.InteropWorkbook.Names;

            foreach (Microsoft.Office.Interop.Excel.Name namedRange in names)
            {
                try
                {
                    Microsoft.Office.Interop.Excel.Range refersToRange = namedRange.RefersToRange;

                    if (refersToRange != null)
                    {
                        ranges.Add(new NamedRange()
                        {
                            WorksheetName = refersToRange.Worksheet.Name,
                            Name = namedRange.Name,
                            Row = refersToRange.Row - 1,
                            Column = refersToRange.Column - 1,
                            Rows = refersToRange.Rows.Count,
                            Columns = refersToRange.Columns.Count,
                        });
                    }
                }
                catch 
                { 
                    // RefersToRange can throw exception
                }
            }

            return ranges;
        }

        public IWorksheet AddWorksheet(int? index, IWorksheet before = null, IWorksheet after = null)
        {
            InteropWorksheet interopWorksheet;

            if (index.HasValue && index.Value == 0)
            {
                interopWorksheet = (InteropWorksheet)this.InteropWorkbook.Sheets.Add();
            }
            else
            {
                if (before != null)
                {
                    interopWorksheet = (InteropWorksheet)this.InteropWorkbook.Sheets.Add(((Worksheet)before).InteropWorksheet);
                }
                else if (after != null)
                {
                    interopWorksheet = (InteropWorksheet)this.InteropWorkbook.Sheets.Add(null, ((Worksheet)after).InteropWorksheet);
                }
                else
                {
                    var sortedWorksheets = this.worksheetByInteropWorksheet.OrderBy(v => v.Value.Index).Select(v => v.Key).ToArray();
                    InteropWorksheet append = null;
                    if (sortedWorksheets.Any())
                    {
                        if (!index.HasValue || index > sortedWorksheets.Length - 1)
                        {
                            index = sortedWorksheets.Length - 1;
                        }

                        append = sortedWorksheets[index.Value];
                    }

                    interopWorksheet = (InteropWorksheet)this.InteropWorkbook.Sheets.Add(Missing.Value, append, Missing.Value, Missing.Value);
                }
            }

            return this.worksheetByInteropWorksheet[interopWorksheet];
        }

        public IWorksheet[] Worksheets => this.worksheetByInteropWorksheet.Values.Cast<IWorksheet>().ToArray();

        public bool IsActive { get; internal set; }

        public void Close(bool? saveChanges = null, string fileName = null)
        {
            this.InteropWorkbook.Close((object)saveChanges ?? Missing.Value, (object)fileName ?? Missing.Value, Missing.Value);
        }

        public Worksheet New(InteropWorksheet interopWorksheet)
        {
            var worksheet = new Worksheet(this, interopWorksheet);
            this.worksheetByInteropWorksheet.Add(interopWorksheet, worksheet);
            
            return worksheet;
        }

        private void Workbook_Activate()
        {
            throw new NotImplementedException();
        }

        private async void ApplicationOnWorkbookNewSheet(InteropWorkbook wb, object sh)
        {
            if (sh is InteropWorksheet interopWorksheet)
            {
                if (!this.worksheetByInteropWorksheet.TryGetValue(interopWorksheet, out var worksheet))
                {
                    worksheet = new Worksheet(this, interopWorksheet);
                    this.worksheetByInteropWorksheet.Add(interopWorksheet, worksheet);
                }

                interopWorksheet.BeforeDelete += async () => await this.AddIn.Program.OnBeforeDelete(worksheet);

                await this.AddIn.Program.OnNew(worksheet);
            }
            else
            {
                Console.WriteLine("Not a InteropWorksheet");
            }
        }

        private void ApplicationOnSheetBeforeDelete(object sh)
        {
            if (sh is InteropWorksheet interopWorksheet)
            {
                this.worksheetByInteropWorksheet.Remove(interopWorksheet);
            }
            else
            {
                Console.WriteLine("Not a InteropWorksheet");
            }
        }
    }
}

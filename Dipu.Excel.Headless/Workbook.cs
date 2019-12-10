using System.Collections.Generic;
using System.Linq;
using System.Transactions;

namespace Dipu.Excel.Headless
{
    public class Workbook : IWorkbook
    {
        public Workbook(AddIn addIn)
        {
            this.AddIn = addIn;
            this.WorksheetList = new List<Worksheet>();
        }

        public AddIn AddIn { get; }

        public List<Worksheet> WorksheetList { get; set; }

        public IWorksheet[] Worksheets => this.WorksheetList.Cast<IWorksheet>().ToArray();

        public bool Active { get; set; }

        public IWorksheet CreateSheet(int? index = null, IWorksheet before = null, IWorksheet after = null)
        {
            var sheet = new Worksheet(this);

            if (index != null)
            {
                this.WorksheetList.Insert(index.Value, sheet);
            }
            else if (before != null)
            {
                this.WorksheetList.Insert(this.WorksheetList.IndexOf(before as Worksheet), sheet);
            }
            else if (after != null)
            {
                this.WorksheetList.Insert(this.WorksheetList.IndexOf(after as Worksheet) + 1, sheet);
            }
            else
            {
                this.WorksheetList.Add(sheet);
            }

            if (this.WorksheetList.All(v => !v.Active))
            {
                var worksheet = this.WorksheetList.FirstOrDefault();
                if (worksheet != null)
                {
                    worksheet.Active = true;
                }
            }

            return sheet;
        }

        public void Close(bool? saveChanges = null, string fileName = null)
        {
            this.AddIn.Remove(this);
        }
    }
}

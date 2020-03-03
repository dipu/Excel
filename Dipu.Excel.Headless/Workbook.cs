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

        public bool IsActive { get; private set; }

        public IWorksheet AddWorksheet(int? index = null, IWorksheet before = null, IWorksheet after = null)
        {
            var worksheet = new Worksheet(this);

            if (index != null)
            {
                this.WorksheetList.Insert(index.Value, worksheet);
            }
            else if (before != null)
            {
                this.WorksheetList.Insert(this.WorksheetList.IndexOf(before as Worksheet), worksheet);
            }
            else if (after != null)
            {
                this.WorksheetList.Insert(this.WorksheetList.IndexOf(after as Worksheet) + 1, worksheet);
            }
            else
            {
                this.WorksheetList.Add(worksheet);
            }

            worksheet.Activate();

            return worksheet;
        }

        public void Close(bool? saveChanges = null, string fileName = null)
        {
            this.AddIn.Remove(this);
        }

        public void Activate()
        {
            foreach (var workbook in this.AddIn.WorkbookList)
            {
                workbook.IsActive = false;
            }

            this.IsActive = true;
        }

        public List<NamedRange> GetNamedRanges()
        {
            throw new System.NotImplementedException();
        }
    }
}

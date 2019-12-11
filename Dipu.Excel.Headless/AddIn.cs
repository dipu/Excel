using System.Collections.Generic;
using System.Linq;

namespace Dipu.Excel.Headless
{
    public class AddIn : IAddIn
    {
        public AddIn()
        {
            this.WorkbookList = new List<Workbook>();
        }

        public IWorkbook[] Workbooks => this.WorkbookList.Cast<IWorkbook>().ToArray();

        public IList<Workbook> WorkbookList { get; }

        public Workbook AddWorkbook()
        {
            var workbook = new Workbook(this);
            this.WorkbookList.Add(workbook);
            workbook.Activate();
            return workbook;
        }

        public void Remove(Workbook workbook)
        {
            this.WorkbookList.Remove(workbook);
        }
    }
}

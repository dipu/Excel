using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Action = System.Action;

namespace Dipu.Spreadsheet.Local
{
    public class Host : IHost
    {
        private readonly Dictionary<string, Action> handlerByAction;
 
        public Host(Application application)
        {
            this.Application = application;
            this.handlerByAction = new Dictionary<string, Action>();
            this.WorkbookByComWorkbook = new Dictionary<Microsoft.Office.Interop.Excel.Workbook, Workbook>();
        }

        public Application Application { get; }

        public Dictionary<Microsoft.Office.Interop.Excel.Workbook, Workbook> WorkbookByComWorkbook { get; }

        public void Register(string action, Action handler)
        {
            this.handlerByAction[action] = handler;
        }

        public IEnumerable<IWorkbook> Workbooks => this.WorkbookByComWorkbook.Values;

        public void Handle(string action)
        {
            if (this.handlerByAction.TryGetValue(action, out var handler))
            {
                handler();
            }
        }

        public Workbook Activate(Microsoft.Office.Interop.Excel.Workbook comWorkbook)
        {
            if (!this.WorkbookByComWorkbook.TryGetValue(comWorkbook, out var workbook))
            {
                workbook = new Workbook(this, comWorkbook);
                this.WorkbookByComWorkbook.Add(comWorkbook, workbook);
            }

            return workbook;
        }
    }
}

using System.Collections.Generic;
using System.Linq;
using Action = System.Action;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace Dipu.Excel.Embedded
{
    public class AddIn : IAddIn
    {
        private readonly Dictionary<string, Action> handlerByAction;
        private readonly Dictionary<ComWrapper<InteropWorkbook>, Workbook> workbookByInteropWorkbook;

        public AddIn(InteropApplication application)
        {
            this.Application = application;
            this.handlerByAction = new Dictionary<string, Action>();
            this.workbookByInteropWorkbook = new Dictionary<ComWrapper<InteropWorkbook>, Workbook>();
        }

        public InteropApplication Application { get; }

        public IReadOnlyDictionary<ComWrapper<InteropWorkbook>, Workbook> WorkbookByInteropWorkbook => workbookByInteropWorkbook;

        public void Register(string action, Action handler)
        {
            this.handlerByAction[action] = handler;
        }

        public IWorkbook[] Workbooks => this.WorkbookByInteropWorkbook.Values.Cast<IWorkbook>().ToArray();

        public void Handle(string action)
        {
            if (this.handlerByAction.TryGetValue(action, out var handler))
            {
                handler();
            }
        }

        public Workbook New(ComWrapper<InteropWorkbook> interopWorkbook)
        {
            if (!this.workbookByInteropWorkbook.TryGetValue(interopWorkbook, out var workbook))
            {
                workbook = new Workbook(this, interopWorkbook);
                this.workbookByInteropWorkbook.Add(interopWorkbook, workbook);
            }

            return workbook;
        }

        public void Close(ComWrapper<InteropWorkbook> interopWorkbook)
        {
            this.workbookByInteropWorkbook.Remove(interopWorkbook);
        }
    }
}

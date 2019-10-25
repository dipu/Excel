using System.Threading;
using System.Windows.Forms;
using Dipu.Excel.Embedded;
using Application;
using AppEvents_Event = Microsoft.Office.Interop.Excel.AppEvents_Event;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ExcelAddInLocal
{
    public partial class ThisAddIn
    {
        private AddIn addIn;
        private Program program;

        private async void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            SynchronizationContext windowsFormsSynchronizationContext = new WindowsFormsSynchronizationContext();
            SynchronizationContext.SetSynchronizationContext(windowsFormsSynchronizationContext);

            this.addIn = new AddIn(this.Application);
            Globals.Ribbons.Ribbon.AddIn = addIn;
            this.program = new Program();
            await program.OnStart(addIn);

            ((AppEvents_Event)this.Application).NewWorkbook += async wb =>
            {
                var interopWorkbook = Com.Wrap(wb);
                var workbook = addIn.New(interopWorkbook);

                using (var interopWorksheets = Com.Wrap(interopWorkbook.ComObject.Worksheets))
                {
                    for (var i = 1; i <= interopWorksheets.ComObject.Count; i++)
                    {
                        var interopWorksheet = Com.Wrap((InteropWorksheet)interopWorksheets.ComObject[i]);
                        workbook.New(interopWorksheet);
                    }

                    var worksheets = workbook.Worksheets;
                    await this.program.OnNew(workbook);
                    foreach (var worksheet in worksheets)
                    {
                        await program.OnNew(worksheet);
                    }
                }
            };

            void workbookBeforeClose(InteropWorkbook wb, ref bool cancel)
            {
                var interopWorkbook = Com.Wrap(wb);
                if (this.addIn.WorkbookByInteropWorkbook.TryGetValue(interopWorkbook, out var workbook))
                {
                    this.program.OnClose(workbook, ref cancel);
                    if (!cancel)
                    {
                        this.addIn.Close(interopWorkbook);
                    }
                }
            }

            this.Application.WorkbookBeforeClose += workbookBeforeClose;
        }

        private async void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            await program.OnStop();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}

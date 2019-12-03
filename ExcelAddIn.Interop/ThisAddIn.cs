using System.ComponentModel;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Dipu.Excel.Embedded;
using Application;
using Nito.AsyncEx;
using AppEvents_Event = Microsoft.Office.Interop.Excel.AppEvents_Event;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ExcelAddInLocal
{
    public partial class ThisAddIn
    {
        private AddIn addIn;

        private void ThisAddIn_Startup(object sender, System.EventArgs e) => AsyncContext.Run(async () =>
        {
            SynchronizationContext windowsFormsSynchronizationContext = new WindowsFormsSynchronizationContext();
            SynchronizationContext.SetSynchronizationContext(windowsFormsSynchronizationContext);

            var program = new Program();
            this.addIn = new AddIn(this.Application, program);
            this.Ribbon.AddIn = this.addIn;
            await program.OnStart(addIn);
        });

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) => AsyncContext.Run(async () =>
        {
            await this.addIn.Program.OnStop();
        });

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            this.Ribbon = new Ribbon();
            return this.Ribbon;
        }

        public Ribbon Ribbon { get; set; }

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

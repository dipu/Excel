using System;
using Dipu.Excel.Embedded;
using Moq;
using Xunit;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace Dipu.Excel.Tests.Embedded
{
    public class WorkbookTests : IDisposable
    {
        private InteropApplication application;

        public WorkbookTests()
        {
            this.application = new InteropApplication { Visible = true };
        }

        public void Dispose()
        {
            foreach (InteropWorkbook workbook in this.application.Workbooks)
            {
                workbook.Close(false);
            }

            this.application.Quit();
        }

        [Fact]
        public async void OnNew()
        {
            var program = new Mock<IProgram>();
            var addIn = new AddIn(application, program.Object);

            application.Workbooks.Add();

            program.Verify(mock => mock.OnNew(It.IsAny<IWorkbook>()), Times.Once());
        }
    }
}

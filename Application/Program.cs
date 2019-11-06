using System.Collections.Generic;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using Dipu.Excel.Rendering;

namespace Application
{
    using System.Drawing;
    using System.Threading.Tasks;
    using Dipu.Excel;

    public class Program : IProgram
    {
        public IAddIn AddIn { get; private set; }

        public async Task OnStart(IAddIn addIn)
        {
            this.AddIn = addIn;

            this.AddIn.Register(Actions.Dosomething, () => MessageBox.Show("Boom!!!!"));
        }

        public async Task OnStop()
        {
        }

        public async Task OnNew(IWorkbook workbook)
        {
            var sheet = workbook.CreateSheet();
            sheet.Name = "2";
            
            var cells = new List<CellValue>();
            for (var i = 0; i < 10 * 1000; i++)
            {
                for (var j = 0; j < 300; j++)
                {
                    cells.Add(new CellValue(i, j, $"'{i},{j}"));
                }
            }

            sheet.Render(cells);
        }

        public void OnClose(IWorkbook workbook, ref bool cancel)
        {
        }

        public async Task OnNew(IWorksheet worksheet)
        {
            worksheet.Name = "1";
        }
    }
}
using System.Linq;
using System.Windows.Forms;

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
            this.CanWriteCellStyle = new Style(Color.LightBlue, Color.Black);
            this.CanNotWriteCellStyle = new Style(Color.MistyRose, Color.Black);
            this.ChangedCellStyle = new Style(Color.DeepSkyBlue, Color.Black);

            var sheet = workbook.CreateSheet();
            sheet.Name = $"{workbook.Worksheets.Length}";

            for (var i = 0; i < 50; i++)
            {
                for (var j = 0; j < 10; j++)
                {
                    sheet[i, j].Value = decimal.Parse($"{i},{j}");
                    if (j == 0 || j == 2)
                    {
                        sheet[i, j].Style = this.CanWriteCellStyle;
                        sheet[i, j].NumberFormat = "#.###,00";
                    }
                    else
                    {
                        sheet[i, j].Style = this.CanNotWriteCellStyle;
                    }
                }
            }

            await sheet.Flush();
            
            sheet[0, 0].Value = "Whoppa!";
            sheet[0, 0].Comment = "De Poppa!";

            sheet[5, 5].Style = this.CanNotWriteCellStyle;

            await sheet.Flush();

            sheet.CellChanged += (sender, v) =>
            {
                foreach (var cell in v.Cells)
                {
                    cell.Style = this.ChangedCellStyle;
                }

                ((IWorksheet) sender).Flush();

                //MessageBox.Show($"Cells changed: {string.Join(",", v.Cells.Select(w => $"{w.Row}:{w.Column}"))}");
            };
        }

        public Style CanNotWriteCellStyle { get; set; }

        public Style CanWriteCellStyle { get; set; }

        public Style ChangedCellStyle { get; set; }

        public void OnClose(IWorkbook workbook, ref bool cancel)
        {
        }

        public async Task OnNew(IWorksheet worksheet)
        {
            worksheet.Name = "1";
        }

        public Task OnBeforeDelete(IWorksheet worksheet)
        {
            return Task.CompletedTask;
        }
    }
}
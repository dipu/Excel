using System.Windows.Forms;
using Microsoft.Extensions.DependencyInjection;

namespace Application
{
    using System.Drawing;
    using System.Threading.Tasks;
    using Dipu.Excel;

    public class Program : IProgram
    {
        public Program(ServiceProvider serviceProvider)
        {
            this.ServiceProvider = serviceProvider;
        }

        public ServiceProvider ServiceProvider { get; }

        public IAddIn AddIn { get; private set; }

        public async Task OnStart(IAddIn addIn)
        {
            this.AddIn = addIn;
        }

        public async Task OnHandle(string handle, params object[] argument)
        {
            switch (handle)
            {
                case Actions.Dosomething:
                    MessageBox.Show("Boom!!!!");
                    break;
            }
        }

        public async Task OnStop()
        {
        }

        public async Task OnNew(IWorkbook workbook)
        {
            this.CanWriteCellStyle = new Style(Color.LightBlue, Color.Black);
            this.CanNotWriteCellStyle = new Style(Color.MistyRose, Color.Black);
            this.ChangedCellStyle = new Style(Color.DeepSkyBlue, Color.Black);

            var sheet = workbook.AddWorksheet();
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

            var style = new Style(Color.Red, Color.White);
            sheet[3, 3].Style = style;
            sheet[3, 5].Style = style;
            sheet[4, 4].Style = style;
            sheet[5, 3].Style = style;
            sheet[5, 5].Style = style;

            await sheet.Flush();

            sheet[0, 0].Value = "Whoppa!";
            sheet[0, 0].Comment = "De Poppa!";

            sheet[10, 2].Style = this.CanNotWriteCellStyle;

            sheet[3, 12].Value = "Walter";
            sheet[3, 13].Value = "Martien";
            sheet[3, 14].Value = "Koen";

            sheet[0, 12].Options = new Range(row: 2, column: 12, columns: 3);

            await sheet.Flush();

            sheet.CellsChanged += (sender, v) =>
            {
                foreach (var cell in v.Cells)
                {
                    cell.Style = this.ChangedCellStyle;
                }

                ((IWorksheet)sender).Flush();

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

        public async Task OnLogin()
        {
        }

        public async Task OnLogout()
        {
        }

        public bool IsEnabled(string controlId, string controlTag)
        {
            return true;
        }
    }
}

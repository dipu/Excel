using System.Windows.Forms;

namespace Application
{
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
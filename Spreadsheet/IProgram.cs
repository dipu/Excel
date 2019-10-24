namespace Dipu.Spreadsheet
{
    using System.Threading.Tasks;

    public interface IProgram
    {
        Task OnStart(IHost host);

        Task OnActivate(IWorkbook workbook);

        Task OnStop();
    }
}

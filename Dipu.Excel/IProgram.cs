namespace Dipu.Excel
{
    using System.Threading.Tasks;

    public interface IProgram
    {
        Task OnStart(IAddIn addIn);

        Task OnStop();

        Task OnNew(IWorkbook workbook);

        void OnClose(IWorkbook workbook, ref bool cancel);
        
        Task OnNew(IWorksheet worksheet);
    }
}

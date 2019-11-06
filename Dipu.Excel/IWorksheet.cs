using System.Collections.Generic;
using Dipu.Excel.Rendering;

namespace Dipu.Excel
{
    public interface IWorksheet
    {
        string Name { get; set; }

        void Render(IEnumerable<CellValue> cellValues);
    }
}

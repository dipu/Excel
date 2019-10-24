using System;
using System.Collections.Generic;

namespace Dipu.Spreadsheet
{
    public interface IWorkbook
    {
        ISheet CreateSheet();

        IEnumerable<ISheet> Sheets { get; }
    }
}

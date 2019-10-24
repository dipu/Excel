using System;
using System.Collections.Generic;

namespace Dipu.Spreadsheet
{
    public interface IHost
    {
        void Register(string action, Action handler);

        IEnumerable<IWorkbook> Workbooks { get; }
    }
}

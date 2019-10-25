using System;

namespace Dipu.Excel
{
    public interface IAddIn
    {
        void Register(string action, Action handler);

        IWorkbook[] Workbooks { get; }
    }
}

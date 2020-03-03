using System.Collections.Generic;

namespace Dipu.Excel
{
    public struct NamedRange
    {
        public string WorksheetName { get; set; }

        public string Name { get; set; }

        public int Row { get; set; }

        public int Column { get; set; }

        public int Rows { get; set; }

        public int Columns { get; set; }      
    }
}

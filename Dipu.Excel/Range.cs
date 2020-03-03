﻿using System;
﻿using System.Collections.Generic;

namespace Dipu.Excel
{
    public class Range
    {
        public Range(int row, int column, int? rows = null, int? columns = null, string name = null)
        {
            this.Row = row;
            this.Column = column;
            this.Rows = rows;
            this.Columns = columns;
            this.Name = name;

            if(this.Columns == null && this.Rows == null)
            {
                throw new ArgumentException("Either Columns or Rows is required.");
            }
        }

        public string Name { get; }

        public int Row { get; }

        public int Column { get; }

        public int? Rows { get; }

        public int? Columns { get; }
    }
}

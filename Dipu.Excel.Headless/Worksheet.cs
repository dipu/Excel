using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Dipu.Excel.Headless
{
    public class Worksheet : IWorksheet
    {
        public Workbook Workbook { get; }

        public Worksheet(Workbook workbook)
        {
            Workbook = workbook;
            this.CellByRowColumn = new Dictionary<string, Cell>();
        }

        public event EventHandler<CellChangedEvent> CellsChanged;

        public string Name { get; set; }

        public bool Active { get; set; }
        
        public Dictionary<string, Cell> CellByRowColumn { get; }

        public ICell this[int row, int column]
        {
            get
            {
                var key = $"{row}:{column}";
                if (!this.CellByRowColumn.TryGetValue(key, out var cell))
                {
                    cell = new Cell(this, row, column);
                    this.CellByRowColumn.Add(key, cell);
                }

                return cell;
            }
        }

        public async Task Flush()
        {
        }
    }
}

using System.Collections.Generic;
using System.Linq;
using Dipu.Excel.Rendering;
using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace Dipu.Excel.Embedded
{
    public class Worksheet : IWorksheet
    {
        public Worksheet(Workbook workbook, InteropWorksheet interopWorksheet)
        {
            this.Workbook = workbook;
            this.InteropWorksheet = interopWorksheet;
        }

        public Workbook Workbook { get; set; }

        public InteropWorksheet InteropWorksheet { get; set; }

        public string Name { get => this.InteropWorksheet.Name; set => this.InteropWorksheet.Name = value; }

        public void Render(IEnumerable<CellValue> cellValues)
        {
            foreach (var chunk in cellValues.Chunks())
            {
                var values = new object[chunk.Count, chunk[0].Count];
                for (var i = 0; i < chunk.Count; i++)
                {
                    for (var j = 0; j < chunk[0].Count; j++)
                    {
                        values[i, j] = chunk[i][j].Value;
                    }
                }

                var fromRow = chunk.First().First().Row;
                var fromColumn = chunk.First().First().Column;

                var toRow = chunk.Last().Last().Row;
                var toColumn = chunk.Last().Last().Column;

                var from = this.InteropWorksheet.Cells[fromRow + 1, fromColumn + 1];
                var to = this.InteropWorksheet.Cells[toRow + 1, toColumn + 1];
                var range = this.InteropWorksheet.Range[from, to];
                range.Value2 = values;
            }
        }

        public int Index => this.InteropWorksheet.Index;
    }
}

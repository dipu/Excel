using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace Dipu.Excel.Embedded
{
    public interface IEmbeddedWorksheet : IWorksheet
    {
        void AddDirtyValue(Cell cell);

        void AddDirtyComment(Cell cell);

        void AddDirtyStyle(Cell cell);

        void AddDirtyNumberFormat(Cell cell);
    }

    public class Worksheet : IEmbeddedWorksheet
    {
        private Dictionary<string, Cell> CellByRowColumn { get; }

        private HashSet<Cell> DirtyValueCells { get; set; }

        private HashSet<Cell> DirtyCommentCells { get; set; }

        private HashSet<Cell> DirtyStyleCells { get; set; }

        private HashSet<Cell> DirtyNumberFormatCells { get; set; }

        public Worksheet(Workbook workbook, InteropWorksheet interopWorksheet)
        {
            this.Workbook = workbook;
            this.InteropWorksheet = interopWorksheet;
            this.CellByRowColumn = new Dictionary<string, Cell>();
            this.DirtyValueCells = new HashSet<Cell>();
            this.DirtyCommentCells = new HashSet<Cell>();
            this.DirtyStyleCells = new HashSet<Cell>();
            this.DirtyNumberFormatCells = new HashSet<Cell>();

            interopWorksheet.Change += InteropWorksheet_Change;

            ((Microsoft.Office.Interop.Excel.DocEvents_Event)interopWorksheet).Activate += () => this.Active = true;
            ((Microsoft.Office.Interop.Excel.DocEvents_Event)interopWorksheet).Deactivate += () => this.Active = false;
        }

        public bool Active { get; private set; }

        private void InteropWorksheet_Change(Range target)
        {
            List<Cell> cells = null;
            foreach (Range targetCell in target.Cells)
            {
                var row = targetCell.Row - 1;
                var column = targetCell.Column - 1;
                var cell = (Cell)this[row, column];

                if (cell.UpdateValue(targetCell.Value2))
                {
                    if (cells == null)
                    {
                        cells = new List<Cell>();
                    }

                    cells.Add(cell);
                }
            }

            if (cells != null)
            {
                this.CellsChanged?.Invoke(this, new CellChangedEvent(cells.Cast<ICell>().ToArray()));
            }
        }

        public Workbook Workbook { get; set; }

        public InteropWorksheet InteropWorksheet { get; set; }

        public string Name { get => this.InteropWorksheet.Name; set => this.InteropWorksheet.Name = value; }

        public event EventHandler<CellChangedEvent> CellsChanged;

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
            this.RenderNumberFormat(this.DirtyNumberFormatCells);
            this.DirtyNumberFormatCells = new HashSet<Cell>();

            this.RenderValue(this.DirtyValueCells);
            this.DirtyValueCells = new HashSet<Cell>();

            this.RenderComments(this.DirtyCommentCells);
            this.DirtyCommentCells = new HashSet<Cell>();

            this.RenderStyle(this.DirtyStyleCells);
            this.DirtyStyleCells = new HashSet<Cell>();
        }
        
        public void RenderValue(IEnumerable<Cell> cells)
        {
            foreach (var chunk in cells.Chunks((v, w) => true))
            {
                try
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
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
        }

        public void RenderComments(IEnumerable<Cell> cells)
        {
            foreach (var cell in cells)
            {
                var partCell = (Range)this.InteropWorksheet.Cells[cell.Row + 1, cell.Column + 1];

                if (partCell.Comment == null)
                {
                    partCell.AddComment();
                    partCell.Comment.Shape.TextFrame.AutoSize = true;
                }

                partCell.Comment.Text(cell.Comment);
            }
        }

        public void RenderStyle(IEnumerable<Cell> cells)
        {
            foreach (var chunk in cells.Chunks((v, w) => Equals(v.Style, w.Style)))
            {
                var fromRow = chunk.First().First().Row;
                var fromColumn = chunk.First().First().Column;

                var toRow = chunk.Last().Last().Row;
                var toColumn = chunk.Last().Last().Column;

                var from = this.InteropWorksheet.Cells[fromRow + 1, fromColumn + 1];
                var to = this.InteropWorksheet.Cells[toRow + 1, toColumn + 1];
                var range = this.InteropWorksheet.Range[from, to];

                var cc = chunk[0][0];
                if (cc.Style != null)
                {
                    range.Interior.Color =  ColorTranslator.ToOle(chunk[0][0].Style.BackgroundColor);
                }
                else
                {
                    range.Interior.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
                }
            }
        }

        public void RenderNumberFormat(IEnumerable<Cell> cells)
        {
            foreach (var chunk in cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat)))
            {
                var fromRow = chunk.First().First().Row;
                var fromColumn = chunk.First().First().Column;

                var toRow = chunk.Last().Last().Row;
                var toColumn = chunk.Last().Last().Column;

                var from = this.InteropWorksheet.Cells[fromRow + 1, fromColumn + 1];
                var to = this.InteropWorksheet.Cells[toRow + 1, toColumn + 1];
                var range = this.InteropWorksheet.Range[from, to];

                range.NumberFormat = chunk[0][0].NumberFormat;
            }
        }

        public int Index => this.InteropWorksheet.Index;

        public void AddDirtyNumberFormat(Cell cell)
        {
            this.DirtyNumberFormatCells.Add(cell);
        }

        public void AddDirtyValue(Cell cell)
        {
            this.DirtyValueCells.Add(cell);
        }

        public void AddDirtyComment(Cell cell)
        {
            this.DirtyCommentCells.Add(cell);
        }

        public void AddDirtyStyle(Cell cell)
        {
            this.DirtyStyleCells.Add(cell);
        }
    }
}

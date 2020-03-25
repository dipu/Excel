// <copyright file="Worksheet.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Dipu.Excel.Embedded
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Office.Interop.Excel;
    using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

    public interface IEmbeddedWorksheet : IWorksheet
    {
        void AddDirtyValue(Cell cell);

        void AddDirtyComment(Cell cell);

        void AddDirtyStyle(Cell cell);

        void AddDirtyNumberFormat(Cell cell);

        void AddDirtyOptions(Cell cell);

        void AddDirtyRow(Row row);
    }

    public class Worksheet : IEmbeddedWorksheet
    {
        private readonly Dictionary<int, Row> rowByIndex;

        private readonly Dictionary<int, Column> columnByIndex;

        public Worksheet(Workbook workbook, InteropWorksheet interopWorksheet)
        {
            this.Workbook = workbook;
            this.InteropWorksheet = interopWorksheet;
            this.rowByIndex = new Dictionary<int, Row>();
            this.columnByIndex = new Dictionary<int, Column>();
            this.CellByRowColumn = new Dictionary<string, Cell>();
            this.DirtyValueCells = new HashSet<Cell>();
            this.DirtyCommentCells = new HashSet<Cell>();
            this.DirtyStyleCells = new HashSet<Cell>();
            this.DirtyOptionCells = new HashSet<Cell>();
            this.DirtyNumberFormatCells = new HashSet<Cell>();
            this.DirtyRows = new HashSet<Row>();

            interopWorksheet.Change += this.InteropWorksheet_Change;

            ((DocEvents_Event)interopWorksheet).Activate += () =>
            {
                this.IsActive = true;
                this.SheetActivated?.Invoke(this, this.Name);
            };

            ((DocEvents_Event)interopWorksheet).Deactivate += () => this.IsActive = false;
        }

        public event EventHandler<CellChangedEvent> CellsChanged;

        public event EventHandler<string> SheetActivated;

        public int Index => this.InteropWorksheet.Index;

        public bool IsActive { get; private set; }

        public Workbook Workbook { get; set; }

        public InteropWorksheet InteropWorksheet { get; set; }

        public string Name { get => this.InteropWorksheet.Name; set => this.InteropWorksheet.Name = value; }

        IWorkbook IWorksheet.Workbook => this.Workbook;

        private Dictionary<string, Cell> CellByRowColumn { get; }

        private HashSet<Cell> DirtyValueCells { get; set; }

        private HashSet<Cell> DirtyCommentCells { get; set; }

        private HashSet<Cell> DirtyStyleCells { get; set; }

        private HashSet<Cell> DirtyOptionCells { get; set; }

        private HashSet<Cell> DirtyNumberFormatCells { get; set; }

        private HashSet<Row> DirtyRows { get; set; }

        public ICell this[int row, int column]
        {
            get
            {
                var key = $"{row}:{column}";
                if (!this.CellByRowColumn.TryGetValue(key, out var cell))
                {
                    cell = new Cell(this, this.Row(row), this.Column(column));
                    this.CellByRowColumn.Add(key, cell);
                }

                return cell;
            }
        }

        public static string ExcelColumnFromNumber(int column)
        {
            string columnString = string.Empty;
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }

            return columnString;
        }

        IRow IWorksheet.Row(int index) => this.Row(index);

        IColumn IWorksheet.Column(int index) => this.Column(index);

        public Row Row(int index)
        {
            if (index < 0)
            {
                throw new ArgumentException("Index can not be negative", nameof(this.Row));
            }

            if (!this.rowByIndex.TryGetValue(index, out var row))
            {
                row = new Row(this, index);
                this.rowByIndex.Add(index, row);
            }

            return row;
        }

        public Column Column(int index)
        {
            if (index < 0)
            {
                throw new ArgumentException(nameof(this.Column));
            }

            if (!this.columnByIndex.TryGetValue(index, out var column))
            {
                column = new Column(this, index);
                this.columnByIndex.Add(index, column);
            }

            return column;
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

            this.SetOptions(this.DirtyOptionCells);
            this.DirtyOptionCells = new HashSet<Cell>();

            this.UpdateRows(this.DirtyRows);
            this.DirtyRows = new HashSet<Row>();

            await Task.CompletedTask;
        }

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

        public void AddDirtyOptions(Cell cell)
        {
            this.DirtyOptionCells.Add(cell);
        }

        public void AddDirtyRow(Row row)
        {
            this.DirtyRows.Add(row);
        }

        private void InteropWorksheet_Change(Range target)
        {
            List<Cell> cells = null;
            foreach (Microsoft.Office.Interop.Excel.Range targetCell in target.Cells)
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

        private void RenderValue(IEnumerable<Cell> cells)
        {
            var chunks = cells.Chunks((v, w) => true);

            Parallel.ForEach(
                chunks,
                chunk =>
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

                    var from = this.InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                    var to = this.InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
                    var range = this.InteropWorksheet.Range[from, to];
                    range.Value2 = values;
                });
        }

        private void RenderComments(IEnumerable<Cell> cells)
        {
            Parallel.ForEach(
                cells,
                cell =>
                {
                    var partCell = (Range)this.InteropWorksheet.Cells[cell.Row.Index + 1, cell.Column.Index + 1];

                    if (partCell.Comment == null)
                    {
                        var comment = partCell.AddComment(cell.Comment);
                        comment.Shape.TextFrame.AutoSize = true;
                    }
                    else
                    {
                        partCell.Comment.Text(cell.Comment);
                    }
                });
        }

        private void RenderStyle(IEnumerable<Cell> cells)
        {
            var chunks = cells.Chunks((v, w) => Equals(v.Style, w.Style));

            Parallel.ForEach(
                chunks,
                chunk =>
                {
                    var fromRow = chunk.First().First().Row;
                    var fromColumn = chunk.First().First().Column;

                    var toRow = chunk.Last().Last().Row;
                    var toColumn = chunk.Last().Last().Column;

                    var from = this.InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                    var to = this.InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
                    var range = this.InteropWorksheet.Range[from, to];

                    var cc = chunk[0][0];
                    if (cc.Style != null)
                    {
                        range.Interior.Color = ColorTranslator.ToOle(chunk[0][0].Style.BackgroundColor);
                    }
                    else
                    {
                        range.Interior.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
                    }
                });
        }

        private void RenderNumberFormat(IEnumerable<Cell> cells)
        {
            var chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat));

            Parallel.ForEach(
                chunks,
                chunk =>
                {
                    var fromRow = chunk.First().First().Row;
                    var fromColumn = chunk.First().First().Column;

                    var toRow = chunk.Last().Last().Row;
                    var toColumn = chunk.Last().Last().Column;

                    var from = this.InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                    var to = this.InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
                    var range = this.InteropWorksheet.Range[from, to];

                    range.NumberFormat = chunk[0][0].NumberFormat;
                });
        }

        private void SetOptions(IEnumerable<Cell> cells)
        {
            var chunks = cells.Chunks((v, w) => Equals(v.Options, w.Options));

            Parallel.ForEach(
                chunks,
                chunk =>
                {
                    var fromRow = chunk.First().First().Row;
                    var fromColumn = chunk.First().First().Column;

                    var toRow = chunk.Last().Last().Row;
                    var toColumn = chunk.Last().Last().Column;

                    var from = this.InteropWorksheet.Cells[fromRow.Index + 1, fromColumn.Index + 1];
                    var to = this.InteropWorksheet.Cells[toRow.Index + 1, toColumn.Index + 1];
                    var range = this.InteropWorksheet.Range[from, to];

                    var cc = chunk[0][0];
                    if (cc.Options != null)
                    {
                        var validationRange = cc.Options.Name;
                        if (string.IsNullOrEmpty(validationRange))
                        {
                            if (cc.Options.Columns.HasValue)
                            {
                                validationRange = $"{cc.Options.Worksheet.Name}!${ExcelColumnFromNumber(cc.Options.Column + 1)}${cc.Options.Row + 1}:${ExcelColumnFromNumber(cc.Options.Column + cc.Options.Columns.Value)}${cc.Options.Row + 1}";
                            }
                            else if (cc.Options.Rows.HasValue)
                            {
                                validationRange = $"{cc.Options.Worksheet.Name}!${ExcelColumnFromNumber(cc.Options.Column + 1)}${cc.Options.Row + 1}:${ExcelColumnFromNumber(cc.Options.Column + 1)}${cc.Options.Row + cc.Options.Rows}";
                            }
                        }

                        try
                        {
                            range.Validation.Delete();
                        }
                        catch (Exception)
                        {
                        }

                        range.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, Type.Missing, $"={validationRange}", Type.Missing);
                        range.Validation.IgnoreBlank = !cc.IsRequired;
                        range.Validation.InCellDropdown = !cc.HideInCellDropdown;
                    }
                    else
                    {
                        try
                        {
                            range.Validation.Delete();
                        }
                        catch (Exception)
                        {
                        }
                    }
                });
        }

        private void UpdateRows(HashSet<Row> dirtyRows)
        {
            var chunks = dirtyRows.OrderBy(w => w.Index).Aggregate(
                        new List<IList<Row>> { new List<Row>() },
                        (acc, w) =>
                        {
                            var list = acc[acc.Count - 1];
                            if (list.Count == 0 || (list[list.Count - 1].Hidden == w.Hidden))
                            {
                                list.Add(w);
                            }
                            else
                            {
                                list = new List<Row> { w };
                                acc.Add(list);
                            }

                            return acc;
                        });

            var updateChunks = chunks.Where(v => v.Count > 0);

            Parallel.ForEach(
                updateChunks,
                chunk =>
                {
                    var fromChunk = chunk.First();
                    var toChunk = chunk.Last();
                    var hidden = fromChunk.Hidden;

                    string from = $"$A${fromChunk.Index + 1}";
                    string to = $"$A${toChunk.Index + 1}";

                    var range = this.InteropWorksheet.Range[from, to];
                    range.EntireRow.Hidden = hidden;
                });
        }
    }
}

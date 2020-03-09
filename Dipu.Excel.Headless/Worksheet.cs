using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Dipu.Excel.Headless
{
    public class Worksheet : IWorksheet
    {
        private readonly Dictionary<int, Row> rowByIndex;

        private readonly Dictionary<int, Column> columnByIndex;

        IWorkbook IWorksheet.Workbook => this.Workbook;

        public Workbook Workbook { get; }

        public Worksheet(Workbook workbook)
        {
            Workbook = workbook;

            this.rowByIndex = new Dictionary<int, Row>();
            this.columnByIndex = new Dictionary<int, Column>();
            this.CellByRowColumn = new Dictionary<string, Cell>();
        }

        public event EventHandler<CellChangedEvent> CellsChanged;
        public event EventHandler<string> SheetActivated;

        public string Name { get; set; }

        public bool IsActive { get; private set; }
        
        public Dictionary<string, Cell> CellByRowColumn { get; }

        public ICell this[int row, int column]
        {
            get
            {
                var key = $"{row}:{column}";
                if (!this.CellByRowColumn.TryGetValue(key, out var cell))
                {
                    cell = new Cell(this, Row(row), Column(column));
                    this.CellByRowColumn.Add(key, cell);
                }

                return cell;
            }
        }

        IRow IWorksheet.Row(int index) => this.Row(index);

        public Row Row(int index)
        {
            if(index < 0)
            {
                throw new ArgumentException("Index can not be negative", nameof(Row));
            }

            if(!this.rowByIndex.TryGetValue(index, out var row))
            {
                row = new Row(this, index);
                this.rowByIndex.Add(index, row);
            }

            return row;
        }

        IColumn IWorksheet.Column(int index) => this.Column(index);

        public Column Column(int index)
        {
            if (index < 0)
            {
                throw new ArgumentException(nameof(Column));
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
        }

        public void Activate()
        {
            foreach (var worksheet in this.Workbook.WorksheetList)
            {
                worksheet.IsActive = false;
            }

            this.IsActive = true;
        }
    }
}

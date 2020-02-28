﻿// <copyright file="Client.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace Dipu.Excel
{
    public partial class Binder
    {
        public IWorksheet Worksheet { get; }

        public event EventHandler ToDomained;

        private IDictionary<ICell, IBinding> bindingByCell = new ConcurrentDictionary<ICell, IBinding>();

        private IList<ICell> boundCells = new List<ICell>();

        private IList<ICell> bindingCells = new List<ICell>();

        public readonly Style changedStyle;

        private readonly IDictionary<ICell, Style> changedCells;

        public Binder(IWorksheet worksheet, Style changedStyle = null)
        {
            this.Worksheet = worksheet;
            this.Worksheet.CellsChanged += Worksheet_CellsChanged;
 
            this.changedStyle = changedStyle;
            if (this.changedStyle != null)
            {
                this.changedCells = new Dictionary<ICell, Style>();
            }
        }

        public void Set(int row, int column, IBinding binding)
        {
            this.Set(this.Worksheet[row, column], binding);
        }

        public void Set(ICell cell, IBinding binding)
        {
            this.bindingByCell[cell] = binding;
            this.bindingCells.Add(cell);
        }

        public ICell[] ToCells()
        {
            var obsoleteCells = this.boundCells.Except(this.bindingCells).ToArray();
            this.boundCells = this.bindingCells;
            this.bindingCells = new List<ICell>();

            foreach (var obsoleteCell in obsoleteCells)
            {
                this.bindingByCell.Remove(obsoleteCell);
            }

            foreach (var kvp in this.bindingByCell)
            {
                var cell = kvp.Key;
                var binding = kvp.Value;
                binding.ToCell(cell);
            }

            return obsoleteCells;
        }

        private void Worksheet_CellsChanged(object sender, CellChangedEvent e)
        {
            foreach (var cell in e.Cells)
            {
                if (this.bindingByCell.TryGetValue(cell, out var binding))
                {
                    if (binding.TwoWayBinding)
                    {
                        binding.ToDomain(cell);

                        if (this.changedStyle != null)
                        {
                            if (!this.changedCells.ContainsKey(cell))
                            {
                                this.changedCells.Add(cell, cell.Style);
                            }
                            cell.Style = this.changedStyle;
                        }
                    }
                    else
                    {
                        binding.ToCell(cell);
                    }
                }
            }

            ToDomained?.Invoke(this, EventArgs.Empty);
        }

        public void ResetChangedCells()
        {
            if (this.changedStyle != null)
            {
                foreach (var kvp in this.changedCells)
                {
                    var cell = kvp.Key;
                    var style = kvp.Value;
                    cell.Style = style;
                }

                this.changedCells.Clear();
            }
        }
    }
}
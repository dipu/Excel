// <copyright file="Client.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;

namespace Dipu.Excel
{
    public partial class Binder
    {
        public IWorksheet Worksheet { get; }

        public IDictionary<ICell, IBinding> BindingByCell { get; } = new ConcurrentDictionary<ICell, IBinding>();

        public event EventHandler ToDomained;

        private IDictionary<ICell, Style> ChangedCells;
        public Binder(IWorksheet worksheet)
        {
            this.Worksheet = worksheet;
            this.Worksheet.CellsChanged += Worksheet_CellsChanged;

            this.ChangedCells = new Dictionary<ICell, Style>();
            this.ChangedStyle = new Style(Color.DeepSkyBlue, Color.Black);
        }

        public Style ChangedStyle { get; set; }

        public void ToCells()
        {
            foreach (var kvp in this.BindingByCell)
            {
                var cell = kvp.Key;
                var binding = kvp.Value;
                binding.ToCell(cell);
            }
        }

        private void Worksheet_CellsChanged(object sender, CellChangedEvent e)
        {
            foreach (var cell in e.Cells)
            {
                if (this.BindingByCell.TryGetValue(cell, out var binding))
                {
                    if (binding.TwoWayBinding)
                    {
                        binding.ToDomain(cell);
                        
                        this.ChangedCells.Add(cell, cell.Style);

                        cell.Style = this.ChangedStyle;
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
            foreach (var kvp in this.ChangedCells)
            {
                var cell = kvp.Key;
                var style = kvp.Value;
                cell.Style = style;
            }

            this.ChangedCells = new Dictionary<ICell, Style>();
        }
    }
}
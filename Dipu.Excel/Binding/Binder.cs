// <copyright file="Client.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace Dipu.Excel
{
    public partial class Binder
    {
        public IWorksheet Worksheet { get; }

        public IDictionary<ICell, IBinding> BindingByCell { get; } = new ConcurrentDictionary<ICell, IBinding>();

        public event EventHandler ToDomained;

        public Binder(IWorksheet worksheet)
        {
            this.Worksheet = worksheet;
            this.Worksheet.CellsChanged += Worksheet_CellsChanged;
        }

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
                    binding.ToDomain(cell);
                }
            }

            ToDomained?.Invoke(this, EventArgs.Empty);
        }
    }
}
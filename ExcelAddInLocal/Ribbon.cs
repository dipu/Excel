using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Dipu.Excel.Embedded;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddInLocal
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs ribbonUiEventArgs)
        {
            this.doSomethingButton.Click += (obj, e) => this.AddIn?.Handle("doSomething");
        }

        public AddIn AddIn { get; set; }
    }
}

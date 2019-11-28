﻿using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Action = System.Action;
using InteropApplication = Microsoft.Office.Interop.Excel.Application;
using InteropWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using InteropWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace Dipu.Excel.Embedded
{
    public class AddIn : IAddIn
    {
        private readonly Dictionary<InteropWorkbook, Workbook> workbookByInteropWorkbook;

        public AddIn(InteropApplication application, IProgram program)
        {
            this.Application = application;
            this.Program = program;
            this.workbookByInteropWorkbook = new Dictionary<InteropWorkbook, Workbook>();

            ((AppEvents_Event)this.Application).NewWorkbook += async interopWorkbook =>
            {
                var workbook = this.New(interopWorkbook);
                for (var i = 1; i <= interopWorkbook.Worksheets.Count; i++)
                {
                    var interopWorksheet = (InteropWorksheet)interopWorkbook.Worksheets[i];
                    workbook.New(interopWorksheet);
                }

                var worksheets = workbook.Worksheets;
                await this.Program.OnNew(workbook);
                foreach (var worksheet in worksheets)
                {
                    await program.OnNew(worksheet);
                }
            };

            void WorkbookBeforeClose(InteropWorkbook interopWorkbook, ref bool cancel)
            {
                if (this.WorkbookByInteropWorkbook.TryGetValue(interopWorkbook, out var workbook))
                {
                    this.Program.OnClose(workbook, ref cancel);
                    if (!cancel)
                    {
                        this.Close(interopWorkbook);
                    }
                }
            }

            this.Application.WorkbookActivate += wb =>
            {
                this.WorkbookByInteropWorkbook[wb].Active = true;
            };

            this.Application.WorkbookDeactivate += wb =>
            {
                // Could already be gone by the WorkbookBeforeClose event
                if (this.WorkbookByInteropWorkbook.TryGetValue(wb, out var workbook))
                {
                    this.WorkbookByInteropWorkbook[wb].Active = false;
                }
            };

            this.Application.WorkbookBeforeClose += WorkbookBeforeClose;
        }

        public InteropApplication Application { get; }

        public IProgram Program { get; }

        public IReadOnlyDictionary<InteropWorkbook, Workbook> WorkbookByInteropWorkbook => workbookByInteropWorkbook;
        
        public IWorkbook[] Workbooks => this.WorkbookByInteropWorkbook.Values.Cast<IWorkbook>().ToArray();
       
        public Workbook New(InteropWorkbook interopWorkbook)
        {
            if (!this.workbookByInteropWorkbook.TryGetValue(interopWorkbook, out var workbook))
            {
                workbook = new Workbook(this, interopWorkbook);
                this.workbookByInteropWorkbook.Add(interopWorkbook, workbook);
            }

            return workbook;
        }

        public void Close(InteropWorkbook interopWorkbook)
        {
            this.workbookByInteropWorkbook.Remove(interopWorkbook);
        }
    }
}

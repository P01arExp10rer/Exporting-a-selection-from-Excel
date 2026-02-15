using System;
using System.Collections.Generic;
using System.Linq;

namespace Exporting_a_selection_from_Excel.Workflow
{
    internal enum SymbolMatchMode
    {
        Exact = 0,
        Contains = 1,
    }

    internal sealed class ExcelSearchSettings
    {
        public string ExcelFilePath { get; set; }
        public string SheetName { get; set; }
        public string SymbolColumnName { get; set; }
        public List<string> Symbols { get; set; } = new List<string>();
        public SymbolMatchMode MatchMode { get; set; } = SymbolMatchMode.Exact;

        public bool PrintToCommandLine { get; set; } = true;
        public bool PlaceMTextInDrawing { get; set; } = true;
        public int MaxRows { get; set; } = 200;

        public void Normalize()
        {
            ExcelFilePath = (ExcelFilePath ?? string.Empty).Trim();
            SheetName = (SheetName ?? string.Empty).Trim();
            SymbolColumnName = (SymbolColumnName ?? string.Empty).Trim();
            Symbols = (Symbols ?? new List<string>())
                .Select(s => (s ?? string.Empty).Trim())
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (MaxRows <= 0) MaxRows = 200;
        }
    }
}

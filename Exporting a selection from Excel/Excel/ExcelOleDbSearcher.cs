using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Exporting_a_selection_from_Excel.Workflow;
using OfficeOpenXml;

namespace Exporting_a_selection_from_Excel.Excel
{
    /// <summary>
    /// Поиск по Excel с помощью EPPlus. Оставляем тот же контракт (DataTable),
    /// чтобы не менять остальной код.
    /// </summary>
    internal static class ExcelOleDbSearcher
    {
        public static DataTable Search(ExcelSearchSettings settings)
        {
            if (settings == null) throw new ArgumentNullException(nameof(settings));
            settings.Normalize();

            if (string.IsNullOrWhiteSpace(settings.ExcelFilePath))
                throw new ArgumentException("Excel file path is required.", nameof(settings.ExcelFilePath));
            if (string.IsNullOrWhiteSpace(settings.SheetName))
                throw new ArgumentException("Sheet name is required.", nameof(settings.SheetName));
            if (string.IsNullOrWhiteSpace(settings.SymbolColumnName))
                throw new ArgumentException("Symbol column name is required.", nameof(settings.SymbolColumnName));
            if (settings.Symbols == null || settings.Symbols.Count == 0)
                throw new ArgumentException("At least one symbol is required.", nameof(settings.Symbols));

            var filePath = settings.ExcelFilePath;
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel file not found.", filePath);

            // EPPlus лицензия (для некоммерческого использования).
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var wb = package.Workbook;
                if (wb == null || wb.Worksheets.Count == 0)
                    throw new InvalidOperationException("Excel workbook has no worksheets.");

                var ws = wb.Worksheets
                    .FirstOrDefault(w => string.Equals(w.Name, settings.SheetName, StringComparison.OrdinalIgnoreCase));

                if (ws == null)
                    throw new ArgumentException($"Sheet '{settings.SheetName}' not found in workbook.", nameof(settings.SheetName));

                if (ws.Dimension == null)
                    return new DataTable(); // пустой лист

                var startRow = ws.Dimension.Start.Row;
                var startCol = ws.Dimension.Start.Column;
                var endRow = ws.Dimension.End.Row;
                var endCol = ws.Dimension.End.Column;

                // Строим DataTable: первая строка — заголовки.
                var dt = new DataTable();
                var headerRow = startRow;

                var columnIndexByName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                for (int col = startCol; col <= endCol; col++)
                {
                    var header = ws.Cells[headerRow, col].Text?.Trim();
                    if (string.IsNullOrWhiteSpace(header))
                        header = $"Column{col - startCol + 1}";

                    // Гарантируем уникальность имён столбцов.
                    var baseName = header;
                    var name = baseName;
                    int suffix = 1;
                    while (dt.Columns.Contains(name))
                    {
                        name = baseName + "_" + suffix++;
                    }

                    dt.Columns.Add(name);
                    columnIndexByName[name] = col;
                }

                // Находим индекс столбца для поиска по имени заголовка.
                var symbolColumnName = settings.SymbolColumnName.Trim();
                int searchColIndex = -1;

                foreach (var kvp in columnIndexByName)
                {
                    if (string.Equals(kvp.Key, symbolColumnName, StringComparison.OrdinalIgnoreCase))
                    {
                        searchColIndex = kvp.Value;
                        break;
                    }
                }

                if (searchColIndex == -1)
                    throw new ArgumentException($"Column '{settings.SymbolColumnName}' not found in header row.", nameof(settings.SymbolColumnName));

                // Подготавливаем список символов для сравнения.
                var symbols = settings.Symbols
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Select(s => s.Trim())
                    .ToList();

                bool isExact = settings.MatchMode == SymbolMatchMode.Exact;

                // Перебираем строки, начиная со следующей после заголовка.
                for (int row = headerRow + 1; row <= endRow; row++)
                {
                    var cellText = ws.Cells[row, searchColIndex].Text ?? string.Empty;
                    var cellNorm = cellText.Trim();

                    if (string.IsNullOrEmpty(cellNorm))
                        continue;

                    bool match = false;

                    if (isExact)
                    {
                        match = symbols.Any(sym =>
                            string.Equals(cellNorm, sym, StringComparison.OrdinalIgnoreCase));
                    }
                    else
                    {
                        match = symbols.Any(sym =>
                            cellNorm.IndexOf(sym, StringComparison.OrdinalIgnoreCase) >= 0);
                    }

                    if (!match)
                        continue;

                    // Копируем всю строку в DataTable.
                    var values = new object[dt.Columns.Count];
                    for (int col = startCol; col <= endCol; col++)
                    {
                        var valueText = ws.Cells[row, col].Text;
                        values[col - startCol] = valueText;
                    }

                    dt.Rows.Add(values);
                }

                return dt;
            }
        }
    }
}


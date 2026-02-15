using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace Exporting_a_selection_from_Excel.Excel
{
    internal static class ExcelSheetNames
    {
        public static List<string> TryGetSheetNames(string excelFilePath)
        {
            if (string.IsNullOrWhiteSpace(excelFilePath) || !File.Exists(excelFilePath))
                return new List<string>();

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var wb = package.Workbook;
                    if (wb == null || wb.Worksheets.Count == 0)
                        return new List<string>();

                    return wb.Worksheets
                        .Select(ws => ws.Name)
                        .Where(n => !string.IsNullOrWhiteSpace(n))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }
            }
            catch
            {
                return new List<string>();
            }
        }
    }
}

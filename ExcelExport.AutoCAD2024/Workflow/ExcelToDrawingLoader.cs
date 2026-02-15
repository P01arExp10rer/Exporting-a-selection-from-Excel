using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace Exporting_a_selection_from_Excel.Workflow
{
    internal static class ExcelToDrawingLoader
    {
        public static void LoadResultIntoDrawing(Document doc, ExcelSearchSettings settings, System.Data.DataTable result)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (settings == null) throw new ArgumentNullException(nameof(settings));
            if (result == null) throw new ArgumentNullException(nameof(result));

            var ed = doc.Editor;

            var rows = result.Rows.Count;
            var cols = result.Columns.Count;

            if (settings.PrintToCommandLine)
            {
                ed.WriteMessage($"\nExcel search: found {rows} row(s), {cols} column(s).");
                ed.WriteMessage($"\nFile: {settings.ExcelFilePath}");
                ed.WriteMessage($"\nSheet: {settings.SheetName}; Symbol column: {settings.SymbolColumnName}; Mode: {settings.MatchMode}");
            }

            if (!settings.PlaceMTextInDrawing)
                return;

            if (rows == 0)
            {
                ed.WriteMessage("\nNothing to place in drawing (0 rows).");
                return;
            }

            var ppo = new PromptPointOptions("\nSpecify insertion point for Excel results MText:");
            var ppr = ed.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK)
                return;

            var max = Math.Max(1, settings.MaxRows);
            var truncated = rows > max;
            var take = Math.Min(rows, max);

            var text = BuildPlainText(result, take, truncated);
            var mtextContents = ToMTextContents(text);

            var db = doc.Database;
            double textHeight = 2.5;
            try
            {
                var v = AcadApp.GetSystemVariable("TEXTSIZE");
                if (v is double d) textHeight = d;
            }
            catch { }

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                var mt = new MText
                {
                    Location = ppr.Value,
                    TextHeight = textHeight,
                    Contents = mtextContents,
                    Attachment = AttachmentPoint.TopLeft,
                };

                btr.AppendEntity(mt);
                tr.AddNewlyCreatedDBObject(mt, true);

                tr.Commit();
            }
        }

        private static string BuildPlainText(System.Data.DataTable dt, int takeRows, bool truncated)
        {
            var colNames = dt.Columns.Cast<System.Data.DataColumn>().Select(c => c.ColumnName).ToList();

            var sb = new StringBuilder();
            sb.AppendLine("Excel search result");
            sb.AppendLine(new string('-', 40));
            sb.AppendLine(string.Join("\t", colNames));

            for (int i = 0; i < takeRows; i++)
            {
                var row = dt.Rows[i];
                var values = new List<string>();
                foreach (var c in dt.Columns.Cast<System.Data.DataColumn>())
                {
                    var v = row[c];
                    values.Add(v == null || v == DBNull.Value ? "" : Convert.ToString(v));
                }
                sb.AppendLine(string.Join("\t", values));
            }

            if (truncated)
                sb.AppendLine($"... truncated (showing first {takeRows} rows) ...");

            return sb.ToString();
        }

        private static string ToMTextContents(string plain)
        {
            if (plain == null) return string.Empty;

            var s = plain
                .Replace("\\", "\\\\")
                .Replace("{", "\\{")
                .Replace("}", "\\}");

            s = s.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\\P");
            return s;
        }
    }
}

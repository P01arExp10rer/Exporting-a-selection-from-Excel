using System;
using Autodesk.AutoCAD.Runtime;
using Exporting_a_selection_from_Excel.Excel;
using Exporting_a_selection_from_Excel.UI;
using Exporting_a_selection_from_Excel.Workflow;
using Autodesk.AutoCAD.ApplicationServices;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace Exporting_a_selection_from_Excel
{
    public class ExcelExportCommands
    {
        [CommandMethod("ExcelExport")]
        public static void DoIt()
        {
            try
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                doc.Editor.WriteMessage("\nКоманда загружена. Используйте ExcelExportMenu для поиска в Excel и загрузки результата в чертеж.");
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
            }
        }

        [CommandMethod("ExcelExportMenu")]
        public static void ExcelLaunchMenu()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                using (var form = new ExcelLaunchForm())
                {
                    var dr = AcadApp.ShowModalDialog(form);
                    if (dr != System.Windows.Forms.DialogResult.OK || form.Settings == null)
                        return;

                    var settings = form.Settings;
                    var result = ExcelOleDbSearcher.Search(settings);

                    ExcelToDrawingLoader.LoadResultIntoDrawing(doc, settings, result);
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nExcelExportMenu failed: {ex.Message}");
            }
        }
    }
}

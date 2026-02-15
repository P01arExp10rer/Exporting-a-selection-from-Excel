using System;
using Gssoft.Gscad.Runtime;
using Exporting_a_selection_from_Excel.Excel;
using Exporting_a_selection_from_Excel.UI;
using Exporting_a_selection_from_Excel.Workflow;
using Gssoft.Gscad.ApplicationServices;
using GscadApp = Gssoft.Gscad.ApplicationServices.Application;

[assembly: CommandClass(typeof(Exporting_a_selection_from_Excel.ExcelExportCommands))]

namespace Exporting_a_selection_from_Excel
{
    public class ExcelExportCommands
    {
        [CommandMethod("ExcelExportMenu")]
        public static void ExcelLaunchMenu()
        {
            var doc = GscadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                using (var form = new ExcelLaunchForm())
                {
                    // Prefer CAD-hosted modal dialog if available in this API.
                    var dr = GscadApp.ShowModalDialog(form);
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

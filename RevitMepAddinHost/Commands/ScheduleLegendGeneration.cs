using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitMepAddinHost.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ScheduleLegendGeneration : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiapp = commandData.Application;

                object result = Loader.Call("RevitLogic.EntryPoints.ScheduleLegendGenerationEntry.Run", uiapp);
                string text = result?.ToString() ?? "";

                TaskDialog.Show("圖例生成(搭配明細表)", text);
                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("圖例生成(搭配明細表) - Error", ex.ToString());
                return Result.Failed;
            }
        }
    }
}

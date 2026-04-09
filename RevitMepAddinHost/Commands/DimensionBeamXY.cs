using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitMepAddinHost.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class DimensionBeamXY : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiapp = commandData.Application;

                object result = Loader.Call("RevitLogic.EntryPoints.DimensionBeamXYEntry.Run", uiapp);
                string text = result?.ToString() ?? "";

                TaskDialog.Show("DimensionBeamXY", text);
                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("DimensionBeamXY - Error", ex.ToString());
                return Result.Failed;
            }
        }
    }
}

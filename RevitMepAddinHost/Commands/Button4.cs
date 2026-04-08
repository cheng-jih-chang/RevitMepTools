using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitMepAddinHost.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class Button4 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiapp = commandData.Application;

                object result = Loader.Call("RevitLogic.EntryPoints.Button4Entry.Run", uiapp);
                string text = result?.ToString() ?? "";

                TaskDialog.Show("Button4", text);
                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("Button4 - Error", ex.ToString());
                return Result.Failed;
            }
        }
    }
}

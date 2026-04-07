using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitMepAddinHost.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class Button2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiapp = commandData.Application;

                object result = Loader.Call("RevitMepLogic.EntryPoints.Button2Entry.Run", uiapp);
                string text = result?.ToString() ?? "";

                TaskDialog.Show("Button2", text);
                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("Button2 - Error", ex.ToString());
                return Result.Failed;
            }
        }
    }
}

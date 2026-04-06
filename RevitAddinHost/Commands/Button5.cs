using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitAddinHost.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class Button5 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiapp = commandData.Application;

                object result = Loader.Call("RevitLogic.EntryPoints.Button5Entry.Run", uiapp);
                string text = result?.ToString() ?? "";

                TaskDialog.Show("Button5", text);
                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("Button5 - Error", ex.ToString());
                return Result.Failed;
            }
        }
    }
}

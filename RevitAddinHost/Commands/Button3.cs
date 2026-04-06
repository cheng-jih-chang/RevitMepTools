using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitAddinHost.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class Button3 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiapp = commandData.Application;

                object result = Loader.Call("RevitLogic.EntryPoints.Button3Entry.Run", uiapp);
                string text = result?.ToString() ?? "";

                TaskDialog.Show("Button3", text);
                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("Button3 - Error", ex.ToString());
                return Result.Failed;
            }
        }
    }
}

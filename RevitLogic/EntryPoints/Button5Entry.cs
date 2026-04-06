using Autodesk.Revit.UI;
using RevitLogic.Features.Button5;

namespace RevitLogic.EntryPoints
{
    public static class Button5Entry
    {
        public static string Run(UIApplication uiapp)
        {
            return new Button5Service().Execute(uiapp);
        }
    }
}

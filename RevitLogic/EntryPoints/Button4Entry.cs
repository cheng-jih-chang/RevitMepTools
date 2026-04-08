using Autodesk.Revit.UI;
using RevitLogic.Features;

namespace RevitLogic.EntryPoints
{
    public static class Button4Entry
    {
        public static string Run(UIApplication uiapp)
        {
            return new Button4Service().Execute(uiapp);
        }
    }
}

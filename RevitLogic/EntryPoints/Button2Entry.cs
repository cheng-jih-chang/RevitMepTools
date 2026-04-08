using Autodesk.Revit.UI;
using RevitLogic.Features;

namespace RevitLogic.EntryPoints
{
    public static class Button2Entry
    {
        public static string Run(UIApplication uiapp)
        {
            return new Button2Service().Execute(uiapp);
        }
    }
}

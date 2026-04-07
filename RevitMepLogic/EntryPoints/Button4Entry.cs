using Autodesk.Revit.UI;
using RevitMepLogic.Features;

namespace RevitMepLogic.EntryPoints
{
    public static class Button4Entry
    {
        public static string Run(UIApplication uiapp)
        {
            return new Button4Service().Execute(uiapp);
        }
    }
}

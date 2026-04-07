using Autodesk.Revit.UI;
using RevitMepLogic.Features;

namespace RevitMepLogic.EntryPoints
{
    public static class Button5Entry
    {
        public static string Run(UIApplication uiapp)
        {
            return new Button5Service().Execute(uiapp);
        }
    }
}

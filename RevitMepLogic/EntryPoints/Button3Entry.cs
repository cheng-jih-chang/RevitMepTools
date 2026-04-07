using Autodesk.Revit.UI;
using RevitMepLogic.Features;

namespace RevitMepLogic.EntryPoints
{
    public static class Button3Entry
    {
        public static string Run(UIApplication uiapp)
        {
            return new Button3Service().Execute(uiapp);
        }
    }
}

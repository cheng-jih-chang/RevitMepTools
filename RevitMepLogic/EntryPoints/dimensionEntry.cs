using Autodesk.Revit.UI;
using RevitMepLogic.Features;

namespace RevitMepLogic.EntryPoints
{
    public static class dimensionEntry
    {
        public static string Run(UIApplication uiapp)
        {
            return new dimensionService().Execute(uiapp);
        }
    }
}

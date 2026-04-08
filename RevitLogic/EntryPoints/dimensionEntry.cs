using Autodesk.Revit.UI;
using RevitLogic.Features;

namespace RevitLogic.EntryPoints
{
    public static class dimensionEntry
    {
        public static string Run(UIApplication uiapp)
        {
            return new dimensionService().Execute(uiapp);
        }
    }
}

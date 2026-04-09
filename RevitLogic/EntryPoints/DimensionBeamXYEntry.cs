using Autodesk.Revit.UI;
using RevitLogic.Features;

namespace RevitLogic.EntryPoints
{
    public static class DimensionBeamXYEntry
    {
        public static string Run(UIApplication uiapp)
        {
            return new DimensionBeamXYService().Execute(uiapp);
        }
    }
}

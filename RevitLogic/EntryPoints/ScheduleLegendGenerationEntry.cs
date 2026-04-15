using Autodesk.Revit.UI;
using RevitLogic.Features.ScheduleLegendGeneration;

namespace RevitLogic.EntryPoints
{
    public static class ScheduleLegendGenerationEntry
    {
        public static string Run(UIApplication uiapp)
        {
            return new ScheduleLegendGenerationService().Execute(uiapp);
        }
    }
}

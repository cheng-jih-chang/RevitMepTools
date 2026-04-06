using Autodesk.Revit.UI;
using RevitLogic.Features.Button3;

namespace RevitLogic.EntryPoints
{
    public static class Button3Entry
    {
        public static string Run(UIApplication uiapp)
        {
            return new Button3Service().Execute(uiapp);
        }
    }
}

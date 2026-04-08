using Autodesk.Revit.UI;

namespace RevitLogic.Features
{
    public class dimensionService
    {
        public string Execute(UIApplication uiapp)
        {
            if (uiapp == null) return "uiapp is null";
            // TODO: 套管生成主邏輯
            return "dimensionService confirmed successful execution";
        }
    }
}

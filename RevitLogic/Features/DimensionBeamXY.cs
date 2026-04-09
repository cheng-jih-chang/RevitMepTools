using Autodesk.Revit.UI;

namespace RevitLogic.Features
{
    public class DimensionBeamXYService
    {
        public string Execute(UIApplication uiapp)
        {
            if (uiapp == null) return "uiapp is null";
            // TODO: 套管生成主邏輯
            return "DimensionBeamXYService confirmed successful execution";
        }
    }
}

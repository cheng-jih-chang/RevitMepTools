using Autodesk.Revit.UI;

namespace RevitLogic.Features
{
    public class Button5Service
    {
        public string Execute(UIApplication uiapp)
        {
            if (uiapp == null) return "uiapp is null";
            // TODO: 套管生成主邏輯
            return "BUTTON5 confirmed successful execution";
        }
    }
}

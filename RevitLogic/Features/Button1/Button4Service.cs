using Autodesk.Revit.UI;

namespace RevitLogic.Features.Button4
{
    public class Button4Service
    {
        public string Execute(UIApplication uiapp)
        {
            if (uiapp == null) return "uiapp is null";
            // TODO: 套管生成主邏輯
            return "BUTTON4 confirmed successful execution";
        }
    }
}

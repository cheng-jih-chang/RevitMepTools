using Autodesk.Revit.UI;

namespace RevitLogic.Features.Button2
{
    public class Button2Service
    {
        public string Execute(UIApplication uiapp)
        {
            if (uiapp == null) return "uiapp is null";
            // TODO: 套管生成主邏輯
            return "BUTTON2 confirmed successful execution";
        }
    }
}

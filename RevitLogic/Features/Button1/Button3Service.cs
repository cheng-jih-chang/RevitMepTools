using Autodesk.Revit.UI;

namespace RevitLogic.Features.Button3
{
    public class Button3Service
    {
        public string Execute(UIApplication uiapp)
        {
            if (uiapp == null) return "uiapp is null";
            // TODO: 套管生成主邏輯
            return "BUTTON3 confirmed successful execution";
        }
    }
}

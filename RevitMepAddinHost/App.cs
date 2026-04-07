// RevitMepAddinHost\App.cs
using System.Globalization;
using System.Reflection;
using System.Text;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace RevitMepAddinHost
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            const string tabName = "Cheng's Mep Tools_v0.1.1";
            const string panelName = "annotation_v0.1.1";

            try { application.CreateRibbonTab(tabName); } catch { }

            RibbonPanel panel = application.CreateRibbonPanel(tabName, panelName);

            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            var dimension = new PushButtonData(
                "Dimension", WrapText("Dimension"), assemblyPath,
                "RevitMepAddinHost.Commands.dimension"
            );

            var b2 = new PushButtonData(
                "Btn2", WrapText("Button 2"), assemblyPath,
                "RevitMepAddinHost.Commands.Button2"
            );

            var b3 = new PushButtonData(
                "Btn3", WrapText("Button 3"), assemblyPath,
                "RevitMepAddinHost.Commands.Button3"
            );

            var b4 = new PushButtonData(
                "Btn4", WrapText("Button 4"), assemblyPath,
                "RevitMepAddinHost.Commands.Button4"
            );

            var b5 = new PushButtonData(
                "Btn5", WrapText("Button 5"), assemblyPath,
                "RevitMepAddinHost.Commands.Button5"
            );

            panel.AddItem(dimension);
            panel.AddItem(b2);
            panel.AddItem(b3);
            panel.AddItem(b4);
            panel.AddItem(b5);
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private static string WrapText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            return text.Replace(" ", "\n");
        }
    }
}
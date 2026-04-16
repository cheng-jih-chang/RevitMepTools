using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows;
using WpfControls = System.Windows.Controls;

namespace RevitMepAddinHost.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ScheduleLegendGeneration : IExternalCommand
    {
        private void ShowTextWindow(string title, string text)
        {
            Window window = new Window();
            window.Title = title;
            window.Width = 700;
            window.Height = 450;
            window.WindowStartupLocation = WindowStartupLocation.CenterScreen;

            WpfControls.TextBox textBox = new WpfControls.TextBox();
            textBox.Text = text;
            textBox.IsReadOnly = true;
            textBox.TextWrapping = TextWrapping.Wrap;
            textBox.VerticalScrollBarVisibility = WpfControls.ScrollBarVisibility.Auto;
            textBox.HorizontalScrollBarVisibility = WpfControls.ScrollBarVisibility.Auto;
            textBox.FontFamily = new System.Windows.Media.FontFamily("Consolas");
            textBox.FontSize = 14;
            textBox.AcceptsReturn = true;

            window.Content = textBox;
            window.ShowDialog();
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiapp = commandData.Application;

                object result = Loader.Call("RevitLogic.EntryPoints.ScheduleLegendGenerationEntry.Run", uiapp);
                string text = result?.ToString() ?? "";

                ShowTextWindow("圖例生成(搭配明細表)", text);
                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("圖例生成(搭配明細表) - Error", ex.ToString());
                return Result.Failed;
            }
        }
    }
}

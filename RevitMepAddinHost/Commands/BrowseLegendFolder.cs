using System;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitMepAddinHost.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class BrowseLegendFolder : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Type dialogType = Type.GetType("System.Windows.Forms.FolderBrowserDialog, System.Windows.Forms", false);
            if (dialogType == null)
            {
                message = "FolderBrowserDialog is not available.";
                return Result.Failed;
            }

            object dialog = Activator.CreateInstance(dialogType);
            var disposable = dialog as IDisposable;
            try
            {
                dialogType.GetProperty("Description")?.SetValue(dialog, "Select Legend DWG Root Folder", null);
                dialogType.GetProperty("SelectedPath")?.SetValue(dialog, App.LegendDwgRootFolderText ?? string.Empty, null);
                dialogType.GetProperty("ShowNewFolderButton")?.SetValue(dialog, false, null);

                object result = dialogType.GetMethod("ShowDialog", Type.EmptyTypes)?.Invoke(dialog, null);
                if (!string.Equals(result?.ToString(), "OK", StringComparison.OrdinalIgnoreCase))
                    return Result.Cancelled;

                string selectedPath = dialogType.GetProperty("SelectedPath")?.GetValue(dialog, null) as string;
                if (string.IsNullOrWhiteSpace(selectedPath) || !Directory.Exists(selectedPath))
                    return Result.Cancelled;

                App.SetLegendRootFolder(selectedPath);
                return Result.Succeeded;
            }
            finally
            {
                disposable?.Dispose();
            }
        }
    }
}

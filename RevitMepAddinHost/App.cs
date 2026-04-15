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
        private const string LegendDwgRootFolderPreset1 = @"C:\Users\sunny\NoteSystem\PublicNotes\CivilWorkMaterials\REVIT\元件\00_使用中圖例";
        private const string LegendDwgRootFolderPreset2 = @"D:\PublicNotes\CivilWorkMaterials\REVIT\元件\00_使用中圖例";
        private const string LegendFolderCurrentMember = "LegendFolderCurrentMember";
        private const string LegendFolderPresetMember1 = "LegendFolderPresetMember1";
        private const string LegendFolderPresetMember2 = "LegendFolderPresetMember2";
        private static ComboBox _legendRootComboBox;
        private static ComboBoxMember _legendCurrentMember;

        public static string LegendDwgRootFolderText { get; private set; } = LegendDwgRootFolderPreset1;

        public Result OnStartup(UIControlledApplication application)
        {
            const string tabName = "Cheng's Mep Tools_v0.1.1";
            const string panelName = "annotation_v0.1.1";

            try { application.CreateRibbonTab(tabName); } catch { }

            RibbonPanel panel = application.CreateRibbonPanel(tabName, panelName);

            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            var dimensionBeamXY = new PushButtonData(
                "Dimension Beam XY", WrapText("Dimension Beam XY"), assemblyPath,
                "RevitMepAddinHost.Commands.DimensionBeamXY"
            );

            var ScheduleLegendGeneration = new PushButtonData(
                "ScheduleLegendGeneration", WrapText("Schedule Legend Generation"), assemblyPath,
                "RevitMepAddinHost.Commands.ScheduleLegendGeneration"
            );

            var browseLegendFolder = new PushButtonData(
                "BrowseLegendFolder", WrapText("Browse Legend Folder"), assemblyPath,
                "RevitMepAddinHost.Commands.BrowseLegendFolder"
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

            panel.AddItem(dimensionBeamXY);
            panel.AddItem(ScheduleLegendGeneration);

            var legendRootComboData = new ComboBoxData("LegendDwgRootFolderPresetCombo");
            _legendRootComboBox = panel.AddItem(legendRootComboData) as ComboBox;
            if (_legendRootComboBox != null)
            {
                var current = new ComboBoxMemberData(LegendFolderCurrentMember, LegendDwgRootFolderText);
                current.ToolTip = "Current selected path";
                _legendCurrentMember = _legendRootComboBox.AddItem(current) as ComboBoxMember;

                var preset1 = new ComboBoxMemberData(LegendFolderPresetMember1, LegendDwgRootFolderPreset1);
                preset1.ToolTip = LegendDwgRootFolderPreset1;
                _legendRootComboBox.AddItem(preset1);

                var preset2 = new ComboBoxMemberData(LegendFolderPresetMember2, LegendDwgRootFolderPreset2);
                preset2.ToolTip = LegendDwgRootFolderPreset2;
                _legendRootComboBox.AddItem(preset2);

                _legendRootComboBox.CurrentChanged += OnLegendRootComboCurrentChanged;
                SetLegendRootFolder(LegendDwgRootFolderPreset1);
            }

            panel.AddItem(browseLegendFolder);

            panel.AddItem(b3);
            panel.AddItem(b4);
            panel.AddItem(b5);
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public static void SetLegendRootFolder(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return;

            LegendDwgRootFolderText = path.Trim();
            if (_legendCurrentMember != null)
            {
                _legendCurrentMember.ItemText = LegendDwgRootFolderText;
                _legendCurrentMember.ToolTip = LegendDwgRootFolderText;
            }
            if (_legendRootComboBox != null && _legendCurrentMember != null)
                _legendRootComboBox.Current = _legendCurrentMember;
        }

        private static string WrapText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            return text.Replace(" ", "\n");
        }

        private static void OnLegendRootComboCurrentChanged(object sender, ComboBoxCurrentChangedEventArgs e)
        {
            ComboBoxMember newItem = e.NewValue;
            if (newItem == null || newItem.Name == LegendFolderCurrentMember)
                return;

            if (newItem.Name == LegendFolderPresetMember1)
                SetLegendRootFolder(LegendDwgRootFolderPreset1);
            else if (newItem.Name == LegendFolderPresetMember2)
                SetLegendRootFolder(LegendDwgRootFolderPreset2);
        }
    }
}
using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.HmiUnified;
using Siemens.Engineering.HmiUnified.UI.Screens;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Siemens.Engineering.HmiUnified.UI.ScreenGroup;
using System.Diagnostics;

namespace ShowScripts
{
    public class AddIn : ContextMenuAddIn //Enthält eigentliche Funktionalität des AddIns
    {
        private readonly TiaPortal _tiaPortal;
        private readonly string exeName = "ShowScripts.OpennessExe.ShowScripts.exe";

        public AddIn(TiaPortal tiaPortal) : base("ShowScriptCode") //Definiert den AddIn-Namen
        {
            _tiaPortal = tiaPortal;
        }

        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Export all scripts of HMI - all screens", OnClickExportOverwriteSilent, DisplayStatus);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Export all scripts of HMI", OnClickExportOverwrite, DisplayStatus);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Export all scripts of HMI - continue from last export - all screens", OnClickExportSilent, DisplayStatus);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Export all scripts of HMI - continue from last export", OnClickExport, DisplayStatus);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Import all scripts to HMI", OnClickImport, DisplayStatus);
        }

        private void OnClickExportSilent(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
#if DEBUG
            Debugger.Launch();
#endif
            string args = "export --silent";
            args = StartApplication(menuSelectionProvider, args);
        }
        private void OnClickExportOverwriteSilent(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
#if DEBUG
            Debugger.Launch();
#endif
            string args = "export --overwrite --silent";
            args = StartApplication(menuSelectionProvider, args);
        }
        private void OnClickExport(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
#if DEBUG
            Debugger.Launch();
#endif
            string args = "export";
            args = StartApplication(menuSelectionProvider, args);
        }
        private void OnClickExportOverwrite(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
#if DEBUG
            Debugger.Launch();
#endif
            string args = "export --overwrite";
            args = StartApplication(menuSelectionProvider, args);
        }

        private void OnClickImport(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
#if DEBUG
            Debugger.Launch();
#endif
            string args = "import";
            args = StartApplication(menuSelectionProvider, args);
        }
        private string StartApplication(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider, string args)
        {
            string fileDirectory = _tiaPortal.Projects.FirstOrDefault()?.Path.DirectoryName + "\\UserFiles\\";
            var currentProcess = _tiaPortal.GetCurrentProcess();
            string exePath = WriteApplicationToTempFolder(exeName);
            args += " -t \"";
            List<string> deviceNames = new List<string>();
            foreach (Device device in menuSelectionProvider.GetSelection<Device>())
            {
                foreach (DeviceItem deviceItem in device.DeviceItems)
                {
                    SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                    if (softwareContainer != null && softwareContainer.Software is HmiSoftware) // HmiSoftware means Unified
                    {
                        deviceNames.Add(device.Name);
                    }
                }
            }
            args += string.Join(",", deviceNames) + "\"";
            args += " -P " + "\"" + currentProcess.Id + "\"";
            args += " -A " + "\"" + currentProcess.Path.ToString() + "\"";

            var process = Siemens.Engineering.AddIn.Utilities.Process.Start(exePath, args);

            return args;
        }


        private string WriteApplicationToTempFolder(string exeResource, string[] referenceResources = null)
        {
            string tempDirectory = GetTemporaryDirectory();

            File.WriteAllBytes(Path.Combine(tempDirectory, GetFileNameFromResource(exeResource)), GetResourceStream(exeResource));

            if (referenceResources != null)
            {
                foreach (string resource in referenceResources)
                {
                    File.WriteAllBytes(Path.Combine(tempDirectory, GetFileNameFromResource(resource)), GetResourceStream(resource));
                }
            }

            return Path.Combine(tempDirectory, GetFileNameFromResource(exeResource));
        }        
        private byte[] GetResourceStream(string name)
        {
            BinaryReader streamReader = new BinaryReader(this.GetType().Assembly.GetManifestResourceStream(name));

            return streamReader.ReadBytes((int)streamReader.BaseStream.Length);
        }
        private string GetTemporaryDirectory()
        {
            string tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }
        private string GetFileNameFromResource(string resourceName)
        {
            string resourceStart = "ShowScripts.OpennessExe.";

            return resourceName.Substring(resourceStart.Length);
        }
        
        private MenuStatus DisplayStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            return MenuStatus.Enabled;
        }
    }
}

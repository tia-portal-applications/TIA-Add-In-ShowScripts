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
using System;
using System.Windows.Forms;
using System.Threading;

namespace ShowScripts
{
    public class AddIn : ContextMenuAddIn //Enthält eigentliche Funktionalität des AddIns
    {
        private readonly TiaPortal _tiaPortal;

        public AddIn(TiaPortal tiaPortal) : base("ShowScriptCode") //Definiert den AddIn-Namen
        {
            _tiaPortal = tiaPortal;
        }

        public void ShowMessages(string message, string messageType, string messageInfo = "")
        {
            var messageBox = _tiaPortal.GetMessageBox();
            messageBox.ShowNotification(NotificationIcon.Information, messageType, message, messageInfo);
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
            Work(menuSelectionProvider, false, false, true);
        }
        private void OnClickExportOverwriteSilent(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
#if DEBUG
            Debugger.Launch();
#endif
            Work(menuSelectionProvider, false, true, true);
        }
        private void OnClickExport(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
#if DEBUG
            Debugger.Launch();
#endif
            Work(menuSelectionProvider, false);
        }
        private void OnClickExportOverwrite(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
#if DEBUG
            Debugger.Launch();
#endif
            Work(menuSelectionProvider, false, true);
        }

        private void OnClickImport(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
#if DEBUG
            Debugger.Launch();
#endif
            Work(menuSelectionProvider, true);
        }
        private void Work(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider, bool isImport, bool overwrite = false, bool silent = false)
        {
            var tiaPortalProject = _tiaPortal.Projects.FirstOrDefault();
            using (var exclusiveAccess = _tiaPortal.ExclusiveAccess("ShowScripts Addin V19.22.0 starting..."))
            {
                foreach (Device device in menuSelectionProvider.GetSelection<Device>())
                {
                    foreach (DeviceItem deviceItem in device.DeviceItems)
                    {
                        SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                        if (softwareContainer != null && softwareContainer.Software is HmiSoftware) // HmiSoftware means Unified
                        {
                            string deviceName = device.Name;
                            bool isExportSuccess = false;
                            var screens = GetScreens(tiaPortalProject, deviceName);
                            if (screens == null)
                            {
                                MessageBox.Show(@"Cannot find any HMI with name: " + deviceName + @"
                            If the device exists, but this error occurs, this is a known bug in TIA Portal. Workaround: Copy the HMI and connected PLC to a new created TIA Portal project and run the tool again.
                            Click to close and continue...");
                                continue;
                            }

                            string fileDirectory = tiaPortalProject.Path.DirectoryName + "\\UserFiles\\ShowScripts_" + deviceName + "\\";
                            if (!Directory.Exists(fileDirectory))
                            {
                                Directory.CreateDirectory(fileDirectory);
                            }
                            // Console.WriteLine("ShowScripts path: " + fileDirectory);

                            var worker = new AddScriptsToList(fileDirectory, exclusiveAccess, deviceName);
                            Thread.Sleep(2000); // give the user time to see what is happening. This short sleep does not harm
                            if (!isImport)
                            {
                                worker.ExportScripts(screens, overwrite, silent, exclusiveAccess,ref isExportSuccess, false);
                                // run command to fix scripts with eslint rules
                                var processStartInfo = new ProcessStartInfo();
                                processStartInfo.WorkingDirectory = fileDirectory;
                                processStartInfo.FileName = "cmd.exe";
                                processStartInfo.Arguments = "/C npm run lint";
                                Process proc = Process.Start(processStartInfo);
                                proc.WaitForExit();
                                if (isExportSuccess)
                                {
                                    ShowMessages("The exporting of scripts is completed. \nThe exported files can be found under the user folder,\n" + fileDirectory, "Information");
                                }
                                else
                                {
                                    ShowMessages("Export of scripts is cancelled", "Information");
                                }
                            }
                            else
                            {
                                FileInfo[] Files = new DirectoryInfo(fileDirectory).GetFiles("*.js");
                                if (Files.Length > 0)
                                {
                                    // import scripts
                                    using (var transaction = exclusiveAccess.Transaction(tiaPortalProject, "Import scripts"))
                                    {
                                        worker.ImportScripts(screens, exclusiveAccess);
                                        ShowMessages("Scripts are imported from path:\n" + fileDirectory, "Information" );
                                    }
                                }
                                else
                                {
                                    ShowMessages("Please export scripts before importing.", "Information");
                                    return;
                                }
                                // tiaPortalProject.Save();
                            }
                        }
                    }
                }
            }
        }
        private static IEnumerable<HmiScreen> GetScreens(Project tiaProject, string deviceName)
        {
            var screens = GetScreens(tiaProject.Devices, deviceName);
            if (screens == null)
            {
                screens = GetScreens(tiaProject.DeviceGroups, deviceName);
            }
            return screens;
        }
        private static IEnumerable<HmiScreen> GetScreens(DeviceUserGroupComposition groups, string deviceName)
        {
            foreach (var group in groups)
            {
                var screens = GetScreens(group.Devices, deviceName);
                if (screens != null)
                {
                    return screens;
                }
                screens = GetScreens(group.Groups, deviceName);
                if (screens != null)
                {
                    return screens;
                }
            }
            return null;
        }

        private static IEnumerable<HmiScreen> GetScreens(DeviceComposition devices, string deviceName)
        {
            IEnumerable<HmiScreen> screens = null;
            if (string.IsNullOrEmpty(deviceName))
            {
                foreach (var device in devices)
                {
                    screens = GetScreens(device);
                    if (screens != null)
                    {
                        deviceName = device.Name;
                        return screens;
                    }
                }
            }
            else
            {
                var device = devices.FirstOrDefault(x => x.Name == deviceName);
                if (device != null)
                {
                    return GetScreens(device);
                }
            }
            return null;
        }
        private static IEnumerable<HmiScreen> GetScreens(Device device)
        {
            foreach (DeviceItem deviceItem in device.DeviceItems)
            {
                SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                if (softwareContainer != null && softwareContainer.Software is HmiSoftware)
                {
                    var sw = (softwareContainer.Software as HmiSoftware);
                    var allScreens = sw.Screens.ToList();
                    allScreens.AddRange(ParseGroups(sw.ScreenGroups));
                    return allScreens;
                }
            }
            return null;
        }

        private static IEnumerable<HmiScreen> ParseGroups(HmiScreenGroupComposition parentGroups)
        {
            foreach (var group in parentGroups)
            {
                foreach (var screen in group.Screens)
                {
                    yield return screen;
                }
                foreach (var screen in ParseGroups(group.Groups))
                {
                    yield return screen;
                }
            }
        }

        private MenuStatus DisplayStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            return MenuStatus.Enabled;
        }
    }
}

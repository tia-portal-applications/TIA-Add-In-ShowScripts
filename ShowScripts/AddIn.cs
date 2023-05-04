using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.HmiUnified;
using Siemens.Engineering.HmiUnified.UI.Screens;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Siemens.Engineering.HmiUnified.UI.ScreenGroup;

namespace ShowScripts
{
    public class AddIn : ContextMenuAddIn //Enthält eigentliche Funktionalität des AddIns
    {
        private readonly TiaPortal _tiaPortal;

        public AddIn(TiaPortal tiaPortal) : base("ShowScriptCode") //Definiert den AddIn-Namen
        {
            _tiaPortal = tiaPortal;
        }

        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Export all scripts of HMI", OnClickExport, DisplayStatus);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Import all scripts to HMI", OnClickImport, DisplayStatus);
        }

        private void OnClickExport(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            //Debugger.Launch();
            
            string screenName = ".*";
            List<ScreenDynEvents> screenDynEvenList = new List<ScreenDynEvents>();
            string fileDirectory = _tiaPortal.Projects.FirstOrDefault()?.Path.DirectoryName + "\\UserFiles\\";
            foreach (IEngineeringObject engObj in menuSelectionProvider.GetSelection())
            {
                AddScriptsToList.ExportScripts(GetScreens(engObj as HardwareObject), fileDirectory, screenName);
            }
        }
        private void OnClickImport(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            //Debugger.Launch();
            
            // string screenName = ".*";
            List<ScreenDynEvents> screenDynEvenList = new List<ScreenDynEvents>();
            string fileDirectory = _tiaPortal.Projects.FirstOrDefault()?.Path.DirectoryName + "\\UserFiles\\";
            foreach (IEngineeringObject engObj in menuSelectionProvider.GetSelection())
            {
                AddScriptsToList.ImportScripts(GetScreens(engObj as HardwareObject), fileDirectory);
            }
        }

        private static IEnumerable<HmiScreen> GetScreens(HardwareObject engObj)
        {
            foreach (DeviceItem deviceItem in engObj.DeviceItems)
            {
                SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                if (softwareContainer != null)
                {
                    var sw = softwareContainer.Software as HmiSoftware;
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

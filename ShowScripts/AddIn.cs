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
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Show all scripts of HMI", OnClick, DisplayStatus); //Definiert den Kontextmenüelementname
        }

        private void OnClick(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            //Debugger.Launch();

            var evelist = new List<string>();
            var dynamizationList = new List<string>();
            var csvString = new List<string>();
            string screenName = ".*";
            List<ScreenDynEvents> screenDynEvenList = new List<ScreenDynEvents>();
            foreach (IEngineeringObject engObj in menuSelectionProvider.GetSelection())
            {
                AddScriptsToList.Do(ref dynamizationList, ref evelist, ref screenDynEvenList, GetScreens(engObj as HardwareObject), ref csvString, ref screenName);
            }

            string userFiles = _tiaPortal.Projects.FirstOrDefault()?.Path.DirectoryName + "\\UserFiles\\";
            
            foreach (var screen in screenDynEvenList)
            {
                using (StreamWriter sw = new StreamWriter(userFiles + screen.Name + "_Dynamizations" + ".js"))
                {
                    sw.Write(string.Join(Environment.NewLine, screen.DynList));
                }
                using (StreamWriter sw = new StreamWriter(userFiles + screen.Name + "_Events" + ".js"))
                {
                    sw.Write(string.Join(Environment.NewLine, screen.EventList));
                }
            }

            using (StreamWriter sw = new StreamWriter(userFiles + "ProjectScriptsDyn.js"))
            {
                sw.Write(string.Join(Environment.NewLine, dynamizationList));
            }
            using (StreamWriter sw = new StreamWriter(userFiles + "ProjectScriptsEve.js"))
            {
                sw.Write(string.Join(Environment.NewLine, evelist));
            }
            using (StreamWriter sw = new StreamWriter(userFiles + "ProjectScriptsOverview.csv"))
            {
                sw.Write(string.Join(Environment.NewLine, csvString));
            }
        }
        
        private static HmiScreenComposition GetScreens(HardwareObject engObj)
        {
            foreach (DeviceItem deviceItem in engObj.DeviceItems)
            {
                SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                if (softwareContainer != null)
                {
                    return (softwareContainer.Software as HmiSoftware).Screens;
                }
            }
            return null;
        }
        private MenuStatus DisplayStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            return MenuStatus.Enabled;
        }
    }
}

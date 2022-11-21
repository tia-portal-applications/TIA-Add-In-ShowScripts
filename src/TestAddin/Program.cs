using Siemens.Engineering;
using Siemens.Engineering.HmiUnified;
using Siemens.Engineering.HmiUnified.UI.Screens;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShowScripts
{
    class Program
    {
        static void Main(string[] args)
        {
            string screenName = ".*";
            string whereConditions = "";
            string sets = "";
            string pathArg = "";
            string projectPath = "";
            TiaPortalMode tiaWithUI = TiaPortalMode.WithoutUserInterface;
            bool quitTiaWhenFinished = false;
            bool closeTiaProj = false;
            bool allScreens = false;
            bool jumpToEnd = false;
            // SELECT * FROM Screen_* SET Dynamization.Trigger.Type=4, Dynamization.Trigger.Tags='Refresh_tag' -WHERE Dynamization.Trigger.Type=250
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].ToUpper() == "-ALLSCREENS")
                {
                    allScreens = true;
                }
                else if (args[i].ToUpper() == "-UPDATE")
                {
                    screenName = args[i + 1];
                    if (screenName == "*") {
                        screenName = ".*";
                    }
                }
                else if (args[i].ToUpper() == "-SET")
                {
                    for (int j = i + 1; j < args.Length; j++, i++)
                    {
                        if (args[j].ToUpper() == "-WHERE")
                        {
                            break;
                        }
                        sets += args[j];
                    }
                }
                else if (args[i].ToUpper() == "-WHERE")
                {
                    for (int j = i + 1; j < args.Length; j++, i++)
                    {
                        whereConditions += args[j];
                    }
                }
                // help
                else if (args[i].ToUpper() == "-H" || args[i].ToUpper() == "-?" || args[i].ToUpper() == "?" )
                {
                    if (args.Length >= (i + 1))
                    {
                        Console.WriteLine("You called the help option!");
                        Console.WriteLine("");
                        Console.WriteLine("-ALLSCREENS        all screens of your project will be exported, ");
                        Console.WriteLine("                   without that you can specify your screen in a following input prompt");
                        Console.WriteLine("-C                 closes the TIA Portal project at the end");
                        Console.WriteLine("-D                 specifies the export directory, if unused the UserFiles directory in the TIA project will be taken");
                        Console.WriteLine("                   example: -D C:\\temp");
                        Console.WriteLine("-H                 calls this help, additionally possible by -?");
                        Console.WriteLine("-?                 calls this help, additionally possible by -H");
                        Console.WriteLine("-P                 specifies the TIA project to use, if unused the already opened project will be used");
                        Console.WriteLine("                   example: -P D:\\TiaProjects\\Digi.ap17");
                        Console.WriteLine("-Q                 closes the TIA Portal instance at the end");
                        //Console.WriteLine("-SET               -SET Dynamization.Trigger.Type=4, Dynamization.Trigger.Tags='Refresh_tag' -WHERE Dynamization.Trigger.Type=250");
                        Console.WriteLine("-U                 only if no TIA Portal is startet, it starts TIA Portal with user interface");
                        Console.WriteLine("-UPDATE            specifies a single screen, to only export that screen");
                        Console.WriteLine("                   example: -U Screen_1");
                        //Console.WriteLine("-WHERE             can be used as a filter ");
                        jumpToEnd = true;
                    }
                }
                // close tia project at the end
                else if (args[i].ToUpper() == "-C")
                {
                    closeTiaProj = true;
                }
                // directory to export data
                else if (args[i].ToUpper() == "-D")
                {
                    if (args.Length >= (i + 1) )
                    {
                        pathArg = args[i + 1];
                        i++;
                    }
                }
                // tia project path
                else if (args[i].ToUpper() == "-P")
                {
                    if (args.Length >= (i + 1))
                    {
                        projectPath = args[i + 1];
                        i++;
                    }
                }
                // close tia instance at the end
                else if (args[i].ToUpper() == "-Q")
                {
                    quitTiaWhenFinished = true;
                }
                // open tia with user interface
                else if (args[i].ToUpper() == "-U")
                {
                    tiaWithUI = TiaPortalMode.WithUserInterface;
                }

            }
            if (!jumpToEnd)
            {
                TiaPortal tiaPortal = null;
                try
                {
                    tiaPortal = TiaPortal.GetProcesses()[0].Attach();
                }
                catch
                {
                    Console.WriteLine(String.Format("Could not find started TIA Portal instance, start new one"));
                    try
                    {
                        tiaPortal = new TiaPortal(tiaWithUI);
                    }
                    catch
                    {
                        Console.WriteLine(String.Format("Could not start TIA Portal"));
                    }
                }

                FileInfo tiaProjectPath = null;//new FileInfo(projectPath);
                Project tiaPortalProject = null;
                if (projectPath != "")
                {
                    tiaProjectPath = new FileInfo(projectPath);
                    try
                    {
                        tiaPortalProject = tiaPortal.Projects.OpenWithUpgrade(tiaProjectPath);
                    }
                    catch
                    {
                        Console.WriteLine(String.Format("Could not open project {0}", tiaProjectPath.FullName));
                    }
                }
                else
                {
                    tiaPortalProject = tiaPortal.Projects.First();
                }

                var screens = GetScreens(tiaPortalProject.Devices);

                var eveList = new List<string>();
                var dynList = new List<string>();
                var csvString = new List<string>();
                List<ScreenDynEvents> screenDynEvenList = new List<ScreenDynEvents>();
                string fileDirectory;
                if (pathArg != "")
                {
                    fileDirectory = pathArg;
                }
                else
                {
                    fileDirectory = tiaPortalProject.Path.DirectoryName + "\\UserFiles\\";
                }
                Console.WriteLine("export path: " + fileDirectory);

                if (!Directory.Exists(fileDirectory))
                {
                    Directory.CreateDirectory(fileDirectory);
                }

                if (allScreens)
                {                    
                    foreach (var screen in screens)
                    {
                        screenName = screen.Name;
                        csvString.Clear();
                        AddScriptsToList.Do(ref dynList, ref eveList, ref screenDynEvenList, screens, ref csvString, ref screenName, whereConditions, sets);
                        using (StreamWriter sw = new StreamWriter(fileDirectory + screenName + "_Dynamizations" + ".js"))
                        {
                            sw.Write(string.Join(Environment.NewLine, dynList));
                        }
                        using (StreamWriter sw = new StreamWriter(fileDirectory + screenName + "_Events" + ".js"))
                        {
                            sw.Write(string.Join(Environment.NewLine, eveList));
                        }
                        using (StreamWriter sw = new StreamWriter(fileDirectory + screenName + "_Scripts_Overview" + ".csv"))
                        {
                            sw.Write(string.Join(Environment.NewLine, csvString));
                        }
                    }
                }
                else
                {
                    AddScriptsToList.Do(ref dynList, ref eveList, ref screenDynEvenList, screens, ref csvString, ref screenName, whereConditions, sets);

                    foreach (var screenInList in screenDynEvenList)
                    {
                        if (screenName == ".*")
                        {   //potentially all screens
                            using (StreamWriter sw = new StreamWriter(fileDirectory + screenInList.Name + "_Dynamizations" + ".js"))
                            {
                                sw.Write(string.Join(Environment.NewLine, screenInList.DynList));
                            }
                            using (StreamWriter sw = new StreamWriter(fileDirectory + screenInList.Name + "_Events" + ".js"))
                            {
                                sw.Write(string.Join(Environment.NewLine, screenInList.EventList));
                            }
                            using (StreamWriter sw = new StreamWriter(fileDirectory + tiaPortalProject.Name + "_Scripts_Overview" + ".csv"))
                            {
                                sw.Write(string.Join(Environment.NewLine, csvString));
                            }
                        }
                        else
                        {   //only one screen
                            if (screenName == screenInList.Name)
                            {
                                using (StreamWriter sw = new StreamWriter(fileDirectory + screenInList.Name + "_Dynamizations" + ".js"))
                                {
                                    sw.Write(string.Join(Environment.NewLine, screenInList.DynList));
                                }
                                using (StreamWriter sw = new StreamWriter(fileDirectory + screenInList.Name + "_Events" + ".js"))
                                {
                                    sw.Write(string.Join(Environment.NewLine, screenInList.EventList));
                                }
                                using (StreamWriter sw = new StreamWriter(fileDirectory + screenInList.Name + "_Scripts_Overview" + ".csv"))
                                {
                                    sw.Write(string.Join(Environment.NewLine, csvString));
                                }
                            }
                        }
                    }
                }
                tiaPortalProject.Save();
                //close TIA projekt
                if (closeTiaProj)
                    tiaPortalProject.Close();
                //close TIA Portal
                if (quitTiaWhenFinished)
                {
                    const string tiaProcessName = "Siemens.Automation.Portal";
                    Process[] processes = Process.GetProcessesByName(tiaProcessName);
                    foreach (var tiaProcess in processes)
                    {
                        tiaProcess.Kill();
                    }
                }
            }
            
            //Console.ReadKey();
        }
        private static HmiScreenComposition GetScreens(DeviceComposition engObj)
        {
            foreach (HardwareObject hw in engObj)
            {
                foreach (DeviceItem deviceItem in hw.DeviceItems)
                {
                    SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                    if (softwareContainer != null && softwareContainer.Software is HmiSoftware)
                    {
                        return (softwareContainer.Software as HmiSoftware).Screens;
                    }
                }
            }
            return null;
        }
    }
}

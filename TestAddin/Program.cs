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
using Siemens.Engineering.HmiUnified.UI.ScreenGroup;
using Siemens.Engineering.HmiUnified.UI.Events;
using Siemens.Engineering.HmiUnified.UI.Base;
using Siemens.Engineering.HmiUnified.UI.Dynamization;

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
            bool jumpToEnd = false;
            bool importScripts = false;
            // SELECT * FROM Screen_* SET Dynamization.Trigger.Type=4, Dynamization.Trigger.Tags='Refresh_tag' -WHERE Dynamization.Trigger.Type=250
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].ToUpper() == "-ALLSCREENS")
                {
                    screenName = ".*";
                }
                else if (args[i].ToUpper() == "-IMPORT")
                {
                    importScripts = true;
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
                        Console.WriteLine("-IMPORT            import all scripts of the directory instead of exporting them");
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
                var processes = TiaPortal.GetProcesses();
                if (processes.Count > 0)
                {
                    try
                    {
                        tiaPortal = processes[0].Attach();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }
                }
                else
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
                
                var screens = GetScreens(tiaPortalProject);

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
                if (!Directory.Exists(fileDirectory))
                {
                    Directory.CreateDirectory(fileDirectory);
                }
                //Console.WriteLine("export path: " + fileDirectory);
                
                if (!importScripts)
                {
                    AddScriptsToList.ExportScripts(screens, fileDirectory, screenName, whereConditions, sets);

                    // run command to fix scripts with eslint rules
                    var processStartInfo = new ProcessStartInfo();
                    processStartInfo.WorkingDirectory = fileDirectory;
                    processStartInfo.FileName = "cmd.exe";
                    processStartInfo.Arguments = "/C npm run lint";
                    Process proc = Process.Start(processStartInfo);
                    proc.WaitForExit();
                }
                else
                {
                    // import scripts
                    AddScriptsToList.ImportScripts(screens, fileDirectory);
                }



                // save project and close if needed
                tiaPortalProject.Save();
                //close TIA projekt
                if (closeTiaProj)
                    tiaPortalProject.Close();
                //close TIA Portal
                if (quitTiaWhenFinished)
                {
                    const string tiaProcessName = "Siemens.Automation.Portal";
                    Process[] allprocesses = Process.GetProcessesByName(tiaProcessName);
                    foreach (var tiaProcess in allprocesses)
                    {
                        tiaProcess.Kill();
                    }
                }
            }

            //Console.ReadKey();
        }

        private static IEnumerable<HmiScreen> GetScreens(Project tiaProject)
        {
            var screens = GetScreens(tiaProject.Devices);
            if (screens == null) {
                screens = GetScreens(tiaProject.DeviceGroups);
            }
            return screens;
        }
        private static IEnumerable<HmiScreen> GetScreens(DeviceUserGroupComposition groups)
        {
            foreach (var group in groups)
            {
                var screens = GetScreens(group.Devices);
                if (screens != null)
                {
                    return screens;
                }
                screens = GetScreens(group.Groups);
                if (screens != null)
                {
                    return screens;
                }
            }
            return null;
        }

        private static IEnumerable<HmiScreen> GetScreens(DeviceComposition engObj)
        {
            foreach (HardwareObject hw in engObj)
            {
                foreach (DeviceItem deviceItem in hw.DeviceItems)
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
    }
}
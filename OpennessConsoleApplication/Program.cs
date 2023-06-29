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
using System.Reflection;
using System.Windows;

namespace ShowScripts
{
    class Program
    {
        private static Dictionary<string, string> cmdArgs;
        private static List<string> deviceNames = new List<string>();
        static void Main(string[] args)
        {
#if DEBUG
            Debugger.Launch();
#endif
            cmdArgs = ParseArguments(args); // e.g. export -t "PC-System_1" -P "8876" -A "C:\Program Files\Siemens\Automation\Portal V18\bin\Siemens.Automation.Portal.exe"
            AppDomain.CurrentDomain.AssemblyResolve += AssemblyResolver;
            if (cmdArgs.ContainsKey("-t"))
            {
                deviceNames = cmdArgs["-t"].Split(',').ToList();
            }
            else // if no device names were set, add at least one empty entry to analysze the first device
            {
                deviceNames.Add("");
            }
            Work(cmdArgs.ContainsKey("-P") ? int.Parse(cmdArgs["-P"]) : -1, args.Length > 0 ? (args[0].ToLower() == "import") : false);
        }
        static void Work(int processId = -1, bool isImport = false)
        {
            string screenName = ".*";
            TiaPortal tiaPortal = null;
            var processes = TiaPortal.GetProcesses();
            if (processes.Count > 0)
            {
                try
                {
                    if (processId == -1)  // just take the first opened TIA Portal, if it is not specified
                    {
                        tiaPortal = processes.First().Attach();
                    }
                    else
                    {
                        tiaPortal = processes.First(x => x.Id == processId).Attach();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
            }
            
            Project tiaPortalProject = tiaPortal.Projects.First();
            foreach (var deviceName in deviceNames)
            {

                var screens = GetScreens(tiaPortalProject, deviceName);

                List<ScreenDynEvents> screenDynEvenList = new List<ScreenDynEvents>();
                string fileDirectory = tiaPortalProject.Path.DirectoryName + "\\UserFiles\\" + deviceName + "\\";
                if (!Directory.Exists(fileDirectory))
                {
                    Directory.CreateDirectory(fileDirectory);
                }
                //Console.WriteLine("export path: " + fileDirectory);

                var worker = new AddScriptsToList();
                if (!isImport)
                {
                    worker.ExportScripts(screens, fileDirectory, screenName, deviceName);

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
                    worker.ImportScripts(screens, fileDirectory);
                    tiaPortalProject.Save();
                }
            }
        }

        private static IEnumerable<HmiScreen> GetScreens(Project tiaProject, string deviceName)
        {
            var screens = GetScreens(tiaProject.Devices, deviceName);
            if (screens == null) {
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


        private static Dictionary<string, string> ParseArguments(string[] args)
        {
            Dictionary<string, string> arguments = new Dictionary<string, string>();
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].ToCharArray()[0] == '-')
                {
                    int x = i + 1;
                    if (x < args.Length)
                    {
                        if (args[x].ToCharArray()[0] != '-')
                        {
                            if (arguments.ContainsKey(args[i]) == false)
                            {
                                arguments.Add(args[i], args[x]);
                            }
                        }
                    }
                }
            }
            return arguments;
        }

        public static Assembly AssemblyResolver(object sender, ResolveEventArgs args)
        {
            int index = args.Name.IndexOf(',');
            string dllPathToTry = string.Empty;
            string directory = string.Empty;
            if (index != -1)
            {
                string name = args.Name.Substring(0, index);
                string path = cmdArgs.ContainsKey("-A") ? cmdArgs["-A"] : "C:\\Program Files\\Siemens\\Automation\\Portal V18\\bin\\Siemens.Automation.Portal.exe";
                if (path != null & path != string.Empty)
                {
                    if (name == "Siemens.Engineering")
                    {
                        try
                        {
                            FileInfo exeFileInfo = new FileInfo(path);
                            dllPathToTry = exeFileInfo.Directory + @"\..\PublicAPI\V18\Siemens.Engineering.dll";
                        }
                        catch (System.NullReferenceException e)
                        {
                            MessageBox.Show("Data2Unified cannot start due to an inconsistent TIA installation. Please contact support.",
                                "Error",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
                            System.Environment.Exit(1);
                        }
                    }
                    else if (name == "Siemens.Engineering.Hmi")
                    {
                        try
                        {
                            FileInfo exeFileInfo = new FileInfo(path);
                            dllPathToTry = exeFileInfo.Directory + @"\..\PublicAPI\V18\Siemens.Engineering.Hmi.dll";
                        }
                        catch (System.NullReferenceException e)
                        {
                            MessageBox.Show("Data2Unified cannot start due to an inconsistent TIA installation. Please contact support.",
                                "Error",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
                            System.Environment.Exit(1);
                        }
                    }
                }



                if (dllPathToTry != string.Empty)
                {
                    string assemblyPath = Path.GetFullPath(dllPathToTry);

                    if (File.Exists(assemblyPath))
                    {
                        return Assembly.LoadFrom(assemblyPath);
                    }
                    else
                    {
                        MessageBox.Show("Data2Unified cannot start due to an inconsistent TIA installation. Please contact support.",
                            "Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                        System.Environment.Exit(1);
                    }
                }
            }
            return null;
        }
    }
}
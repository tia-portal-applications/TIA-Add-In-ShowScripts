using Siemens.Engineering;
using Siemens.Engineering.HmiUnified;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.HW;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Siemens.Engineering.HmiUnified.UI.ScreenGroup;
using Siemens.Engineering.HmiUnified.UI.Screens;

namespace UnifiedOpennessLibrary
{
    public class CmdArgument
    {
        /// <summary>
        /// This is how you may access the option afterwards in your code, e.g. if you set it to "MyOption" you may access via CmdArgs["MyOption"]
        /// </summary>
        public string OptionToSet = "";
        /// <summary>
        /// this is how the user can define this option by a shortcut. Recommended is to use one dash as prefix, like this: -m
        /// </summary>
        public string OptionShort = "";
        /// <summary>
        /// this is how the user can define this option by a long name. Recommended is to use two dashes as prefix, like this: --myoption
        /// </summary>
        public string OptionLong = "";
        /// <summary>
        /// the help text for your option will be shown, if the user forgets to add a required option or if he types -h or --help
        /// </summary>
        public string HelpText = "";
        /// <summary>
        /// the default value of your option, e.g. "yes"
        /// </summary>
        public string Default = "";
        /// <summary>
        /// define if your option is required or not. If it is required, the tool will stop, if the user does not add this option. Default: false.
        /// </summary>
        public bool Required = false;
        /// <summary>
        /// internal bool to check, if this option was already set by the user or not and then to check, if all required options are set
        /// </summary>
        internal bool IsParsed = false;
    }
    public class UnifiedOpennessConnector : IDisposable
    {
        public Dictionary<string, string> CmdArgs { get; private set; } = new Dictionary<string, string>();

        private string TiaPortalVersion { get; set; }
        public string FileDirectory { get; private set; }
        public ExclusiveAccess AccessObject { get; private set; }
        public Project TiaPortalProject { get; private set; }
        public TiaPortal TiaPortal { get; private set; }
        public string DeviceName { get; private set; }
        public HmiSoftware UnifiedSoftware { get; private set; }
        public IEnumerable<HmiScreen> Screens { get; private set; }
        /// <summary>
        /// will be generated in the constructor by the currently running TIA Portal process version
        /// </summary>
        private static string opennessDll;

        /// <summary>
        /// If your tool changes anything on the TIA Portal project, please use transactions!
        /// </summary>
        /// <param name="tiaPortalVersion">e.g. V18 or V19. It must be the part of the path in the installation folder and is the version that has been tested by you with your program</param>
        /// <param name="args">just pass the arguments that you got from the command line here. You may have access via the public member "CmdArgs" to your arguments afterwards</param>
        /// <param name="toolName">define the name of the tool (exe), so help text and the waiting text is more beautiful</param>
        /// <param name="additionalParameters"> The following parameters are already there, so be careful with the short option of new parameters
        /// new CmdArgument() { Default = "", OptionToSet = "ProcessId", OptionShort = "-id", OptionLong = "--processid", HelpText = "define a process id the tool connects to. If empty, the first TIA Portal process will be connected to" } ,
        /// new CmdArgument() { Default = "", OptionToSet = "Include", OptionShort = "-i", OptionLong = "--include", HelpText = "add a list of screen names on which the tool will work on, split by semicolon (cannot be combined with --exclude), e.g. \"Screen_1;My screen 2\"" } ,
        /// new CmdArgument() { Default = "", OptionToSet = "Exclude", OptionShort = "-e", OptionLong = "--exclude", HelpText = "add a list of screen names on which the tool will not work on, split by semicolon (cannot be combined with --include), e.g. \"Screen_1;My screen 2\"" },
        /// new CmdArgument() { Default = "", OptionToSet = "ProjectPath", OptionShort = "-p", OptionLong = "--projectpath", HelpText = @"if you have no TIA Portal opened, the tool can open it for you and open the project from this path (ProcessId will be ignored, if this is set), e.g. D:\projects\Project1\Project1.ap18" },
        /// new CmdArgument() { Default = "yes", OptionToSet = "ShowUI", OptionShort = "-ui", OptionLong = "--showui", HelpText = "if you provided a ProjectPath via -p you may decide, if TIA Portal should be opened with GUI or without, e.g. \"yes\" or \"no\"" },
        /// new CmdArgument() { Default = "no", OptionToSet = "ClosingOnExit", OptionShort = "-c", OptionLong = "--closeonexit", HelpText = "you may decide, if the TIA Portal should be saved and closed when this tool is finished, e.g. \"yes\" or \"no\"" }
        /// </param>
        public UnifiedOpennessConnector(string tiaPortalVersion, string[] args, IEnumerable<CmdArgument> additionalParameters, string toolName = "MyTool")
        {
            TiaPortalVersion = tiaPortalVersion;
            ParseArguments(args.ToList(), toolName, additionalParameters);
            var tiaProcesses = System.Diagnostics.Process.GetProcessesByName("Siemens.Automation.Portal");
            if (tiaProcesses.Length == 0)
            {
                throw new Exception("No TIA Portal instance is running. Please start TIA Portal and open a project with a WinCC Unified device and run this app again!");
            }
            string processPath = tiaProcesses[0].MainModule.FileName;
            string tiaPortalDirectory = Path.GetDirectoryName(processPath);
            // 0. Load TIA Openness DLL dynamically, so it works also in all further versions
            opennessDll = tiaPortalDirectory + "\\..\\PublicAPI\\" + TiaPortalVersion + "\\Siemens.Engineering.dll";
            AppDomain.CurrentDomain.AssemblyResolve += AssemblyResolver;
            Work(toolName);
        }
        void Work(string toolName)
        {
            if (!string.IsNullOrWhiteSpace(CmdArgs["ProjectPath"]))
            {
                TiaPortal = new TiaPortal(CmdArgs["ShowUI"] == "yes" ? TiaPortalMode.WithUserInterface : TiaPortalMode.WithoutUserInterface);
                TiaPortalProject = TiaPortal.Projects.Open(new FileInfo(CmdArgs["ProjectPath"]));
            }
            else
            {
                var processes = TiaPortal.GetProcesses();
                if (processes.Count > 0)
                {
                    try
                    {
                        if (CmdArgs["ProcessId"] == "")  // just take the first opened TIA Portal, if it is not specified
                        {
                            TiaPortal = processes.First().Attach();
                        }
                        else
                        {
                            TiaPortal = processes.First(x => x.Id == int.Parse(CmdArgs["ProcessId"])).Attach();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }
                }
                else
                {
                    throw new Exception("No TIA Portal instance is open. Please open TIA Portal with your project.");
                }
                TiaPortalProject = TiaPortal.Projects.FirstOrDefault();
            }
            if (TiaPortalProject == null)
            {
                throw new Exception("Please check, if the search is working in TIA Portal and install missing GSD files if not. Then run this tool again.");
            }
            AccessObject = TiaPortal.ExclusiveAccess(toolName + " tool running.\nBe careful: Cancelling this request will not cancel the tool. Please close the Command line to avoid any changes.");
            FileDirectory = TiaPortalProject.Path.DirectoryName + "\\UserFiles\\" + toolName + "_" + DeviceName + "\\";
            if (!Directory.Exists(FileDirectory))
            {
                Directory.CreateDirectory(FileDirectory);
            }
            SetHmiByDeviceName();
        }

        private void SetHmiByDeviceName()
        {
            var hmiSoftwares = GetHmiSoftwares();
            UnifiedSoftware = hmiSoftwares.FirstOrDefault(x => (x.Parent as SoftwareContainer).OwnedBy.Container.TypeIdentifier.Contains("Rack.PC") ?
                x.Name == DeviceName : (x.Parent as SoftwareContainer).OwnedBy.Container.Name == DeviceName);
            if (UnifiedSoftware == null)
            {
                throw new Exception("Device with name " + DeviceName + " cannot be found. Please check, if the search is working in TIA Portal and install missing GSD files if not. Then run this tool again.");
            }
            Screens = GetScreens();
            if (!string.IsNullOrWhiteSpace(CmdArgs["Include"]))
            {
                var screenNames = CmdArgs["Include"].Split(';').Where(x => !string.IsNullOrWhiteSpace(x));
                Screens = Screens.Where(x => screenNames.Contains(x.Name));
            }
            else if (!string.IsNullOrWhiteSpace(CmdArgs["Exclude"]))
            {
                var screenNames = CmdArgs["Exclude"].Split(';').Where(x => !string.IsNullOrWhiteSpace(x));
                Screens = Screens.Where(x => !screenNames.Contains(x.Name));
            }
        }

        private IEnumerable<HmiScreen> GetScreens()
        {
            var allScreens = UnifiedSoftware.Screens.ToList();
            allScreens.AddRange(ParseGroups(UnifiedSoftware.ScreenGroups));
            return allScreens;
        }

        private IEnumerable<HmiScreen> ParseGroups(HmiScreenGroupComposition parentGroups)
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
        private IEnumerable<Device> GetAllDevices(DeviceUserGroupComposition parentGroups)
        {
            foreach (var parentGroup in parentGroups)
            {
                foreach (var device in parentGroup.Devices)
                {
                    yield return device;
                }
                foreach (var device in GetAllDevices(parentGroup.Groups))
                {
                    yield return device;
                }
            }
        }

        private IEnumerable<HmiSoftware> GetHmiSoftwares()
        {
            return
                from device in TiaPortalProject.Devices.Concat(GetAllDevices(TiaPortalProject.DeviceGroups))
                from deviceItem in device.DeviceItems
                let softwareContainer = deviceItem.GetService<SoftwareContainer>()
                where softwareContainer?.Software is HmiSoftware
                select softwareContainer.Software as HmiSoftware;
        }

        public void Dispose()
        {
            if (CmdArgs["ClosingOnExit"] == "yes")
            {
                TiaPortalProject?.Save();
                TiaPortalProject?.Close();
            }
            AccessObject?.Dispose();
            TiaPortal?.Dispose();
        }

        public Assembly AssemblyResolver(object sender, ResolveEventArgs args)
        {
            int index = args.Name.IndexOf("Siemens.Engineering,");

            if (index != -1 || args.Name == "Siemens.Engineering")
            {
                return Assembly.LoadFrom(opennessDll);
            }

            return null;
        }
        /// <summary>
        /// parses the arguments of the input string, e.g. HMI_RT_1 -p=1234 --include="Screen_1;Screen 5" will be parsed to elements in the dictionairy: -p with string "1234" and --include with string "Screen_1;Screen 5"
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        private void ParseArguments(List<string> args, string toolname, IEnumerable<CmdArgument> additionalParameters)
        {
            var argConfiguration = new List<CmdArgument>() {
                new CmdArgument() { Default = "", OptionToSet = "ProcessId", OptionShort = "-id", OptionLong = "--processid", HelpText = "define a process id the tool connects to. If empty, the first TIA Portal process will be connected to" } ,
                new CmdArgument() { Default = "", OptionToSet = "Include", OptionShort = "-i", OptionLong = "--include", HelpText = "add a list of screen names on which the tool will work on, split by semicolon (cannot be combined with --exclude), e.g. \"Screen_1;My screen 2\"" } ,
                new CmdArgument() { Default = "", OptionToSet = "Exclude", OptionShort = "-e", OptionLong = "--exclude", HelpText = "add a list of screen names on which the tool will not work on, split by semicolon (cannot be combined with --include), e.g. \"Screen_1;My screen 2\"" },
                new CmdArgument() { Default = "", OptionToSet = "ProjectPath", OptionShort = "-p", OptionLong = "--projectpath", HelpText = @"if you have no TIA Portal opened, the tool can open it for you and open the project from this path (ProcessId will be ignored, if this is set), e.g. D:\projects\Project1\Project1.ap18" },
                new CmdArgument() { Default = "yes", OptionToSet = "ShowUI", OptionShort = "-ui", OptionLong = "--showui", HelpText = "if you provided a ProjectPath via -p you may decide, if TIA Portal should be opened with GUI or without, e.g. \"yes\" or \"no\"" },
                new CmdArgument() { Default = "no", OptionToSet = "ClosingOnExit", OptionShort = "-c", OptionLong = "--closeonexit", HelpText = "you may decide, if the TIA Portal should be saved and closed when this tool is finished, e.g. \"yes\" or \"no\"" }
            };
            if (args.Count == 0)
            {
                DisplayHelp(argConfiguration, toolname);
                throw new Exception("There must be at least one argument to define the device name.");
            }
            if (args.Contains("-h") || args.Contains("--help"))
            {
                DisplayHelp(argConfiguration, toolname);
                throw new Exception();
            }
            DeviceName = args[0];
            args.RemoveAt(0);
            if (additionalParameters != null)
            {
                argConfiguration.AddRange(additionalParameters);
            }
            // set default values
            foreach (var cmdArg in argConfiguration)
            {
                CmdArgs[cmdArg.OptionToSet] = cmdArg.Default;
            }
            // set values from command line
            foreach (var arg in args)
            {
                foreach (var cmdArg in argConfiguration)
                {
                    if (SetParameter(arg, cmdArg))
                    {
                        break; // if setting the parameter successfully, go to the next one
                    }
                }
            }
            var notSetRequiredArgs = argConfiguration.Where(x => x.Required && !x.IsParsed).Select(x => x.OptionToSet);
            if (notSetRequiredArgs.Count() > 0)
            {
                DisplayHelp(argConfiguration, toolname);
                throw new Exception("The following arguments must be set, but were not set via command line: " + string.Join(",", notSetRequiredArgs));
            }
        }
        private bool SetParameter(string arg, CmdArgument cmdArg)
        {
            if ((cmdArg.OptionShort != "" && arg.ToLower().StartsWith(cmdArg.OptionShort)) || (cmdArg.OptionLong != "" && arg.ToLower().StartsWith(cmdArg.OptionLong)))
            {
                var parts = arg.Split('=').ToList();
                if (parts.Count == 2)
                {
                    parts.RemoveAt(0);
                    CmdArgs[cmdArg.OptionToSet] = string.Join("=", parts).Trim('"');
                    cmdArg.IsParsed = true;
                    return true;
                }
            }
            return false;
        }
        static void DisplayHelp(List<CmdArgument> argConfiguration, string toolName)
        {
            string helpText = @"
Usage: " + toolName + @".exe DEVICE_NAME [OPTION]

Always add a DEVICE_NAME. This is the name of your device, that you can see inside the 'Project tree', e.g. HMI_1

Options:
";
            foreach (var argConfig in argConfiguration)
            {
                helpText += argConfig.OptionShort + "\t" + argConfig.OptionLong + "\t\t\t" + argConfig.HelpText + "\n\t\t\t\t\t\t(default: " + argConfig.Default + ") IsReqired: " + argConfig.Required + "\n";
            }
            Console.WriteLine(helpText);
        }
    }
}

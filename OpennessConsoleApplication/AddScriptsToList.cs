using Siemens.Engineering;
using Siemens.Engineering.HmiUnified.UI;
using Siemens.Engineering.HmiUnified.UI.Base;
using Siemens.Engineering.HmiUnified.UI.Dynamization.Script;
using Siemens.Engineering.HmiUnified.UI.Screens;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using Siemens.Engineering.HmiUnified.UI.Controls;
using System.Globalization;
using Siemens.Engineering.HmiUnified.UI.Dynamization;
using Siemens.Engineering.HmiUnified.UI.Events;
using System.IO;
using System.Windows.Forms;
using Siemens.Engineering.HmiUnified.UI.Widgets;

namespace ShowScripts
{
    class AddScriptsToList
    {
        List<string> screenItemEvents = new List<string>();
        List<string> screenItemDynamisations = new List<string>();
        List<string> tagNames = new List<string>();
        Dictionary<string, int> countDynTriggers;
        Dictionary<string, int> countScreenItems;
        Dictionary<string, int> faceplateTypes;
        string globalDefinitionAreaScriptCodeDyns;
        string globalDefinitionAreaScriptCodeEvents;
        List<string> childScreens;
        int check = 0;
        string delimiter = ";";

        public void ImportScripts(IEnumerable<HmiScreen> screens, string fileDirectory)
        {
            FileInfo[] Files = new DirectoryInfo(fileDirectory).GetFiles("*.js");
            foreach (HmiScreen screen in screens)
            {
                foreach (FileInfo file in Files)
                {
                    if (file.Name.EndsWith("_Events.js") && file.Name.StartsWith(screen.Name))
                    {
                        string tempPath = fileDirectory + file.Name;
                        string[] lines = File.ReadAllLines(tempPath);
                        foreach (HmiScreenEventHandler eveHandler in screen.EventHandlers)
                        {
                            var script = eveHandler.GetAttribute("Script") as ScriptDynamization;
                            string eveType = eveHandler.GetAttribute("EventType").ToString();
                            SetScriptFromFile(script, lines, eveType, screen.Name);
                        }
                        foreach (PropertyEventHandler evePropHandler in screen.PropertyEventHandlers)
                        {
                            var script = evePropHandler.GetAttribute("Script") as ScriptDynamization;
                            SetScriptFromFile(script, lines, "OnPropertyChanged", screen.Name, evePropHandler.PropertyName);
                        }
                        HmiScreenItemBaseComposition items = screen.ScreenItems;
                        foreach (HmiScreenItemBase item in items)
                        {
                            foreach (PropertyEventHandler evePropHandlerItem in item.PropertyEventHandlers)
                            {
                                var script = evePropHandlerItem.GetAttribute("Script") as ScriptDynamization;
                                SetScriptFromFile(script, lines, "OnPropertyChanged", item.Name, evePropHandlerItem.PropertyName);
                            }
                            foreach (IEngineeringObject eveHandItem in (item as IEngineeringObject).GetComposition("EventHandlers") as IEngineeringComposition)
                            {
                                var script = eveHandItem.GetAttribute("Script") as ScriptDynamization;
                                string eveType = eveHandItem.GetAttribute("EventType").ToString();
                                SetScriptFromFile(script, lines, eveType, item.Name);
                            }
                        }

                    }
                    else if (file.Name.EndsWith("_Dynamizations.js") && file.Name.StartsWith(screen.Name))
                    {
                        string tempPath = fileDirectory + file.Name;
                        string[] lines = System.IO.File.ReadAllLines(tempPath);
                        foreach (DynamizationBase dynamizationScreen in screen.Dynamizations)
                        {
                            SetScriptFromFile(dynamizationScreen as IHmiScript, lines, "Trigger", screen.Name, dynamizationScreen.PropertyName);
                        }
                        HmiScreenItemBaseComposition items = screen.ScreenItems;
                        foreach (HmiScreenItemBase item in items)
                        {
                            foreach (DynamizationBase dynamizationItem in item.Dynamizations.Where(x => x.DynamizationType == DynamizationType.Script))
                            {
                                SetScriptFromFile(dynamizationItem as IHmiScript, lines, "Trigger", item.Name, dynamizationItem.PropertyName);
                            }
                        }
                    }
                }
            }
        }

        private void SetScriptFromFile(IHmiScript script, string[] lines, string eventName, string itemName, string propertyName = "")
        {
            string functionName = "function _" + itemName.Replace(" ", "_") + (propertyName == "" ? "" : "_" + propertyName) + "_" + eventName + "(";
            bool getLines = false;
            List<string> scriptLines = new List<string>();
            foreach (string line in lines)
            {
                if (!getLines) // find first occurrence of the desired script
                {
                    if (line.Contains(functionName))
                    {
                        // take the rest in this line behind the openign bracket of the function to not loose any code
                        scriptLines.Add(line.Substring(line.IndexOf('{') + 1));
                        getLines = true;
                    }
                }
                else // get all the script lines until next script or file end
                {
                    scriptLines.Add(line);
                    if (line.Contains("//eslint-disable-next-line camelcase")) // indicates next script begins
                    {
                        break;
                    }

                }
            }
            if (scriptLines.Count == 0)
            {
                Console.WriteLine("Cannot find function with name " + functionName + " in script file.");
                return;
            }
            int lastBracketIndex = scriptLines.FindLastIndex(x => x.Trim() == "}");
            scriptLines.RemoveRange(lastBracketIndex, scriptLines.Count - lastBracketIndex); // remove the last closing bracket }
            string newScriptCode = string.Join(Environment.NewLine, scriptLines);
            if (script.ScriptCode.Trim() != newScriptCode.Trim())
            {
                script.ScriptCode = newScriptCode;
            }
            else
            {
                Console.WriteLine($"Did not change funtion {functionName}, because the script code is equal");
            }
        }

        /// <summary>
        /// export all scripts to files
        /// </summary>
        /// <param name="list"></param>
        /// <param name="screens"></param>
        /// <param name="csvString"></param>
        /// <param name="deepSearch"></param>
        /// <param name="screenName">Case ignored, e.g. *, Screen_*</param>
        /// <param name="whereCondition">e.g. Dynamization.Trigger.Type=250</param>
        /// <param name="sets">e.g. Dynamization.Trigger.Type=4, Dynamization.Trigger.Tags='Refresh_tag'</param>
        public void ExportScripts(IEnumerable<HmiScreen> screens, string fileDirectory, string screenName, string deviceName, bool overwrite, bool silent, bool deepSearch = false)
        {
            string whereCondition = "";
            string sets = "";
            var csvStringP = new List<string>();
            string _screenName = "";
            
            if(!silent && screenName == ".*" && check == 0)
            {
                InputForm dialog = new InputForm(deviceName + " - Enter a screen name: ");
                dialog.Text = "Screen name";

                // Show testDialog as a modal dialog and determine if DialogResult = OK.
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    // Read the contents of testDialog's TextBox.
                    _screenName = dialog.ScreenName;
                    dialog.Dispose();
                }
                else
                {
                    dialog.Dispose();
                    return;
                }
                check = 1;
            }
            else
            {
                _screenName = screenName;
            }

            /*if (_screenName == "All screens")
            {
                screenName = ".*";
            }
            else if (_screenName == "")
            {
                screenName = "";
            }*/

            if (_screenName != "" && _screenName != ".*" && screenName == ".*")
            {
                screenName = _screenName;
            }

            using (StreamWriter sw = File.CreateText(fileDirectory + "TextboxesWithoutText.csv"))
            {
                sw.WriteLine(string.Format("Screen{0}Object", delimiter));
            }
            using (StreamWriter sw = File.CreateText(fileDirectory + "ScreenItemsOutOfRange.csv"))
            {
                sw.WriteLine(string.Format("Screen{0}Object", delimiter));
            }
            using (StreamWriter sw = new StreamWriter(fileDirectory + "CyclicTrigger.csv"))
            {
                sw.WriteLine(string.Format("Screen{0}Object{0}Event/Dynamization{0}Cyclic trigger", delimiter));
            }
            using (StreamWriter sw = new StreamWriter(fileDirectory + "TagSetUsages.csv"))
            {
                sw.WriteLine(string.Format("Screen{0}Object{0}Event/Dynamization", delimiter));
            }

            csvStringP.Add(string.Format("Screen name{0}Item count{0}Cycles{0}Disabled{0}Tags-Dyn-Scripts{0}Tags-Dyn{0} Total Number of Tags{0}Resource list{0}Events{0}Child screens{0}ScreenItemsOutOfRange{0}TextBoxesWithoutText{0}UseTagSet", delimiter));

            countScreenItems = new Dictionary<string, int>()
                {
                    { "HmiFaceplateContainer", 0 },
                    { "HmiScreenWindow", 0 },
                    { "HmiGraphicView", 0 },
                    { "HmiAlarmControl", 0 },
                    { "HmiMediaControl", 0 },
                    { "HmiTrendControl", 0 },
                    { "HmiTrendCompanion", 0 },
                    { "HmiProcessControl", 0 },
                    { "HmiFunctionTrendControl", 0 },
                    { "HmiWebControl", 0 },
                    { "HmiDetailedParameterControl", 0 },
                    { "HmiButton", 0 },
                    { "HmiIOField", 0 },
                    { "HmiToggleSwitch", 0 },
                    { "HmiCheckBoxGroup", 0 },
                    { "HmiBar", 0 },
                    { "HmiGauge", 0 },
                    { "HmiSlider", 0 },
                    { "HmiRadioButtonGroup", 0 },
                    { "HmiListBox", 0 },
                    { "HmiClock", 0 },
                    { "HmiTextBox", 0 },
                    { "HmiLine", 0 },
                    { "HmiPolyline", 0 },
                    { "HmiPolygon", 0 },
                    { "HmiEllipse", 0 },
                    { "HmiEllipseSegment", 0 },
                    { "HmiCircleSegment", 0 },
                    { "HmiEllipticalArc", 0 },
                    { "HmiCircularArc", 0 },
                    { "HmiCircle", 0 },
                    { "HmiRectangle", 0 },
                    { "HmiTouchArea", 0 },
                    { "HmiSymbolicIOField", 0 },
                    { "HmiCustomWebControlContainer", 0 },
                    { "HmiCustomWidgetContainer", 0 },
                    { "HmiSystemDiagnosisControl", 0 }
                };

            foreach (var entry in countScreenItems)
            {
                csvStringP[0] += delimiter + entry.Key;
            }

            faceplateTypes = new Dictionary<string, int>();
            
            var screensToExport = screens.Where(s => Regex.Matches(s.Name, _screenName, RegexOptions.IgnoreCase).Count > 0).ToList();
            var screenCount = screensToExport.Count;
            // if TIA Portal crashes, you can find out easily, which screen made the crash by checking which files were created and which is the next screen within this list.
            using (StreamWriter sw = new StreamWriter(fileDirectory + "ScreenNames.txt"))
            {
                foreach (var screen in screensToExport)
                {
                    sw.WriteLine(screen.Name);
                }
            }
            for (var i = 0; i < screenCount; i++)
            {
                var screen = screensToExport[i];
                string screenNamePath = string.Join("_", screen.Name.Split(Path.GetInvalidFileNameChars())).Replace(delimiter, "_");
                if (!overwrite && File.Exists(fileDirectory + screenNamePath + "_Dynamizations.js") && File.Exists(fileDirectory + screenNamePath + "_Events.js"))
                {
                    continue;
                }
                var dynamizationList = new List<string>();
                var eveList = new List<string>();
                var screenDynEventList = new List<ScreenDynEvents>();
                // inits
                Console.WriteLine();
                Console.Write("[" + (i + 1) + "/" + screenCount + "]" + screen.Name);
                tagNames = new List<string>();
                List<string> _dynamizationList = new List<string>(); 
                List<string> _eveList = new List<string>();
                int startIndex = dynamizationList.Count;
                countDynTriggers = new Dictionary<string, int>()
                {
                    { "Disabled", 0 },
                    { "OtherCycles", 0 },
                    { "Tags", 0 },
                    { "T10s", 0 },
                    { "T5s", 0 },
                    { "T2s", 0 },
                    { "T1s", 0 },
                    { "T500ms", 0 },
                    { "T250ms", 0 },
                    { "T100ms", 0 },
                    { "ResourceLists", 0 },
                    { "TagDynamizations", 0 }
                };
                globalDefinitionAreaScriptCodeDyns = "";
                globalDefinitionAreaScriptCodeEvents = "";
                childScreens = new List<string>();
                var screenItemsOutOfRange = new List<string>();
                var cycles = new List<string>();
                var countTextboxesWithoutText = new List<string>();
                IEnumerable<string> tagSetUsages = new List<string>();

                foreach (var key in countScreenItems.Keys.ToList())
                {
                    countScreenItems[key] = 0;
                }

                foreach (var key in faceplateTypes.Keys.ToList())
                {
                    faceplateTypes[key] = 0;
                }

                // calculations
                var screenDynsPropEves = GetAllMyAttributesDynPropEves(screen, deepSearch, whereCondition.Split(',').ToList(), sets.Split(',').ToList(), ref tagSetUsages, ref cycles);
                var screenDyns = screenDynsPropEves[0];
                var screenPropEves = screenDynsPropEves[1];
                var screenEves = GetAllMyAttributesEve(screen, whereCondition.Split(',').ToList(), sets.Split(',').ToList(), ref tagSetUsages, ref cycles);
                uint screenWidth = screen.Width;
                uint screenHeight = screen.Height;

                foreach (var screenitem in screen.ScreenItems)
                {
                    Console.Write('.');  // the user wants to see that something happens, so a dot will be printed for every screenitem
                    var screenitemDynsPropEves = GetAllMyAttributesDynPropEves(screenitem, deepSearch, whereCondition.Split(',').ToList(), sets.Split(',').ToList(), ref tagSetUsages, ref cycles);
                    var screenitemDyns = screenitemDynsPropEves[0];
                    var screenitemPropEves = screenitemDynsPropEves[1];
                    var screenitemEves = GetAllMyAttributesEve(screenitem, whereCondition.Split(',').ToList(), sets.Split(',').ToList(), ref tagSetUsages, ref cycles);

                    screenItemEvents = screenItemEvents.Concat(screenitemEves).ToList().Concat(screenitemPropEves).ToList();

                    screenItemDynamisations = screenItemDynamisations.Concat(screenitemDyns).ToList();

                    if (screenitem is HmiScreenWindow)
                    {
                        childScreens.Add((screenitem as HmiScreenWindow).Screen);
                    }
                    else if (screenitem is HmiFaceplateContainer)
                    {
                        var faceplate = screenitem as HmiFaceplateContainer;
                        if (!faceplateTypes.ContainsKey(faceplate.ContainedType))
                        {
                            faceplateTypes.Add(faceplate.ContainedType, 0);
                            csvStringP[0] += delimiter + faceplate.ContainedType;
                        }
                        faceplateTypes[faceplate.ContainedType]++;
                    }
                    if (countScreenItems.ContainsKey(screenitem.GetType().Name))
                    {
                        countScreenItems[screenitem.GetType().Name]++;
                    }
                    else
                    {
                        Console.WriteLine("Screenitem Type: " + screenitem.GetType().Name + " is unknown.");
                    }
                    if (!(screenitem is Siemens.Engineering.HmiUnified.UI.Shapes.HmiCentricShapeBase))
                    {
                        var dimensions = screenitem.GetAttributes(new List<string>() { "Left", "Top", "Width", "Height" });
                        int left = (int)dimensions[0];
                        int top = (int)dimensions[1];
                        uint width = (uint)dimensions[2];
                        uint height = (uint)dimensions[3];
                        if (left + width < 0 || top + height < 0 || left > screenWidth || top > screenHeight)
                        {
                            screenItemsOutOfRange.Add(screenitem.Name);
                        }
                    }
                    if (screenitem is HmiTextBox)
                    {
                        foreach (var item in (screenitem as HmiTextBox).Text.Items)  // check if there is any text of any language that is empty (default texts are usually "Text", so not all textboxes will be catched, if we try to check, if all texts are empty. So we now check, if any text is empty
                        {
                            if (string.IsNullOrWhiteSpace(item.Text.Replace("<body><p>", "").Replace("</p></body>", "")))
                            {
                                countTextboxesWithoutText.Add(screenitem.Name);
                                break;
                            }
                        }
                    }
                }

                var dynList = screenDyns.Concat(screenItemDynamisations).ToList();
                var dynListCount = dynList.Count / 2;//adding eslint ignore doubles the entries                
                CultureInfo ci = new CultureInfo("en-US");
                DateTime actDateTime = DateTime.Now;
                string actDateTimeLong = actDateTime.ToString("F", ci);
                dynamizationList.Add(Environment.NewLine + "// ********** This file was exported on : " + actDateTimeLong + " **********");
                dynamizationList.Add(Environment.NewLine + "// ********** " + screen.Name + " Item count: " + screen.ScreenItems.Count + " **********");
                dynamizationList.Add(Environment.NewLine + "// ********** " + screen.Name + "_Dynamizations (" + dynListCount + ") **********");                               
                dynamizationList.Add(Environment.NewLine + globalDefinitionAreaScriptCodeDyns);
                foreach (var dyn in dynList)
                {
                    dynamizationList.Add(dyn);
                }

                var eventList = screenEves.Concat(screenPropEves).ToList().Concat(screenItemEvents).ToList();
                var eventListCount = eventList.Count / 2;//adding eslint ignore doubles the entries
                eveList.Add(Environment.NewLine + "// ********** This file was exported on : " + actDateTimeLong + " **********");
                eveList.Add(Environment.NewLine + "// **********    " + screen.Name + "_Events (" + eventListCount + ")    **********");
                eveList.Add(Environment.NewLine + globalDefinitionAreaScriptCodeEvents);
                foreach (var eve in eventList)
                {
                    eveList.Add(eve);
                }

                //seperate file per screen
                _dynamizationList.Add(Environment.NewLine + "// ********** This file was exported on : " + actDateTimeLong + " **********");
                _dynamizationList.Add(Environment.NewLine + "// ********** " + screen.Name + " Item count: " + screen.ScreenItems.Count + " **********");                
                _dynamizationList.Add(Environment.NewLine + "// ********** " + screen.Name + "_Dynamizations (" + dynListCount + ") **********");
                _dynamizationList.Add(Environment.NewLine + globalDefinitionAreaScriptCodeDyns);
                foreach (var dyn in dynList)
                {
                    _dynamizationList.Add(dyn);
                }

                _eveList.Add(Environment.NewLine + "// ********** This file was exported on : " + actDateTimeLong + " **********");
                _eveList.Add(Environment.NewLine + "// **********    " + screen.Name + "_Events (" + eventListCount + ")    **********");                
                _eveList.Add(Environment.NewLine + globalDefinitionAreaScriptCodeEvents);
                foreach (var eve in eventList)
                {
                    _eveList.Add(eve);
                }
                ScreenDynEvents screenInfo = new ScreenDynEvents(screen.Name, ref _dynamizationList, ref _eveList);
                screenDynEventList.Add(screenInfo);

                screenItemEvents.Clear();
                screenItemDynamisations.Clear();
                foreach (var item in countDynTriggers)
                {
                    dynamizationList.Insert(startIndex, "// " + item.Key.PadRight(17) + ": " + item.Value.ToString().PadLeft(3));

                }
                dynamizationList.Insert(startIndex, "// **********    Overview: Trigger counts of dynamizations of screen: " + screen.Name + "    **********");
                csvStringP.Add(string.Format(screen.Name + "{0}" + screen.ScreenItems.Count + "{0}" + (countDynTriggers["T100ms"] + countDynTriggers["T250ms"] + countDynTriggers["T500ms"] +
                    countDynTriggers["T1s"] + countDynTriggers["T2s"] + countDynTriggers["T5s"] + countDynTriggers["T10s"] + countDynTriggers["OtherCycles"]) + "{0}" +
                    countDynTriggers["Disabled"] + "{0}" + countDynTriggers["Tags"] + "{0}" + countDynTriggers["TagDynamizations"] + "{0}" + tagNames.Count() + "{0}" + countDynTriggers["ResourceLists"] + "{0}" +
                    eventListCount + "{0}" + string.Join("&", childScreens) + "{0}" + screenItemsOutOfRange.Count + "{0}" + countTextboxesWithoutText.Count + "{0}" + tagSetUsages.Count(), delimiter));

                foreach (var entry in countScreenItems)
                {
                    csvStringP[csvStringP.Count - 1] += delimiter + entry.Value;
                }

                foreach (var entry in faceplateTypes)
                {
                    csvStringP[csvStringP.Count - 1] += delimiter + entry.Value;
                }
                csvStringP[csvStringP.Count - 1] += delimiter;

                if (dynList.Count() > 0)
                {
                    using (StreamWriter sw = new StreamWriter(fileDirectory + screenNamePath + "_Dynamizations.js"))
                    {
                        sw.Write(string.Join(Environment.NewLine, dynList));
                    }
                }
                if (eveList.Count() > 3)
                {
                    using (StreamWriter sw = new StreamWriter(fileDirectory + screenNamePath + "_Events.js"))
                    {
                        sw.Write(string.Join(Environment.NewLine, eveList));
                    }
                }
                if (countTextboxesWithoutText.Count() > 0)
                {
                    using (StreamWriter sw = new StreamWriter(fileDirectory + "TextboxesWithoutText.csv", true))
                    {
                        foreach (var item in countTextboxesWithoutText)
                        {
                            sw.WriteLine(screen.Name + delimiter + item);
                        }
                    }
                }
                using (StreamWriter sw = new StreamWriter(fileDirectory + "ScreenItemsOutOfRange.csv", true))
                {
                    foreach (var item in screenItemsOutOfRange)
                    {
                        sw.WriteLine(screen.Name + delimiter + item);
                    }
                }
                using (StreamWriter sw = new StreamWriter(fileDirectory + "TagSetUsages.csv", true))
                {
                    foreach (var item in tagSetUsages)
                    {
                        sw.WriteLine(screen.Name + delimiter + item);
                    }
                }
                using (StreamWriter sw = new StreamWriter(fileDirectory + "CyclicTrigger.csv", true))
                {
                    foreach (var item in cycles)
                    {
                        sw.WriteLine(screen.Name + delimiter + item);
                    }
                }
            }

            //fill empty space with 0 
            char cDelimiter = delimiter[0];
            int longestEntry = csvStringP[0].Count(x => x == cDelimiter);
            int counterFaceplateTypes = 0;
            foreach (var entry in csvStringP.ToList())
            {
                if (counterFaceplateTypes != 0)
                {
                    for (int i = entry.Count(x => x == cDelimiter); i <= longestEntry; i++)
                    {
                        csvStringP[counterFaceplateTypes] += "0" + delimiter;
                    }
                }
                counterFaceplateTypes++;
            }
            var rtName = "";
            try {
                if (screens.Count() > 0)
                {
                    var parent = screens.First().Parent;
                    while (!(parent is Siemens.Engineering.HmiUnified.HmiSoftware))
                    {
                        parent = parent.Parent;
                    }
                    rtName = (parent as Siemens.Engineering.HmiUnified.HmiSoftware).Name;
                }
            }
            catch (Exception ex)
            {
                rtName = "Dummy_HMI_RT";
                Console.WriteLine("Failed to define the runtime name correctly, so use " + rtName + "Exception:");
                Console.WriteLine(ex);
            }

            using (StreamWriter sw = new StreamWriter(fileDirectory + rtName + "_Scripts_Overview.csv"))
            {
                sw.Write(string.Join(Environment.NewLine, csvStringP));
            }

            //check if files are filled or not and delete if they are not filled
            int lineCount;
            using (StreamReader sr = new StreamReader(fileDirectory + "CyclicTrigger.csv"))
            {
                lineCount = sr.ReadToEnd().Split('\n').Length;
            }
            if (lineCount <= 2) //2 because header and empty line afterwards
            {
                File.Delete(fileDirectory + "CyclicTrigger.csv");
            }

            lineCount = 0;
            using (StreamReader sr = new StreamReader(fileDirectory + "ScreenItemsOutOfRange.csv"))
            {
                lineCount = sr.ReadToEnd().Split('\n').Length;
            }
            if (lineCount <= 2)//2 because header and empty line afterwards
            {
                File.Delete(fileDirectory + "ScreenItemsOutOfRange.csv");
            }

            lineCount = 0;
            using (StreamReader sr = new StreamReader(fileDirectory + "TagSetUsages.csv"))
            {
                lineCount = sr.ReadToEnd().Split('\n').Length;
            }
            if (lineCount <= 2)//2 because header and empty line afterwards
            {
                File.Delete(fileDirectory + "TagSetUsages.csv");
            }

            lineCount = 0;
            using (StreamReader sr = new StreamReader(fileDirectory + "TextboxesWithoutText.csv"))
            {
                lineCount = sr.ReadToEnd().Split('\n').Length;
            }
            if (lineCount <= 2)//2 because header and empty line afterwards
            {
                File.Delete(fileDirectory + "TextboxesWithoutText.csv");
            }
        }

        public List<List<string>> GetAllMyAttributesDynPropEves(IEngineeringObject obj, bool deepSearch, List<string> whereConditions, List<string> sets, ref IEnumerable<string> tagSetUsage, ref List<string> cycles)
        {
            var tempListDyn = new List<string>();
            var tempListPropEve = new List<string>();

            var dynamizations = GetDynamizations(obj, whereConditions, sets, ref cycles);
            var propertyEvents = GetPropertyEvents(obj, whereConditions, sets, ref cycles);

            var objectName = obj.GetAttribute("Name").ToString().Replace(' ', '_').Replace('-', '_').Replace('&', 'ß');
            foreach (var itemsDyn in dynamizations)
            {
                if (itemsDyn.Value.Count == 2 && (obj is HmiScreen || obj is HmiScreenItemBase))
                {
                    Console.Write(';');  // the user wants to see that something happens, so a semicolon will be printed for every script
                    tempListDyn.Insert(0, itemsDyn.Value[0]);
                    string script = itemsDyn.Value[1];
                    tagSetUsage = tagSetUsage.Concat(GetTagSetUsage(script, objectName + delimiter + itemsDyn.Key));
                    tempListDyn.Insert(1, "function _" + objectName + "_" + itemsDyn.Key + "_Trigger() {" + script + Environment.NewLine + "}");
                }

                if (itemsDyn.Value.Count == 1 && (obj is HmiScreen || obj is HmiScreenItemBase))
                {
                    Console.Write(';');  // the user wants to see that something happens, so a semicolon will be printed for every script
                    tempListDyn.Add(Environment.NewLine + "//eslint-disable-next-line camelcase");
                    string script = itemsDyn.Value[0];
                    tagSetUsage = tagSetUsage.Concat(GetTagSetUsage(script, objectName + delimiter + itemsDyn.Key));
                    tempListDyn.Add("function _" + objectName + "_" + itemsDyn.Key + "_Trigger() {" + script + Environment.NewLine + "}");
                }

                if (itemsDyn.Value.Count == 1 && !(obj is HmiScreen || obj is HmiScreenItemBase))
                {
                    Console.Write(';');  // the user wants to see that something happens, so a semicolon will be printed for every script
                    tempListDyn.Add(Environment.NewLine + "//eslint-disable-next-line camelcase");
                    string script = itemsDyn.Value[0];
                    tagSetUsage = tagSetUsage.Concat(GetTagSetUsage(script, itemsDyn.Key));
                    tempListDyn.Add("_" + itemsDyn.Key + "_Trigger() {" + Environment.NewLine + script + Environment.NewLine + "}");
                }
            }

            foreach (var itemsPropEve in propertyEvents)
            {
                Console.Write(';');  // the user wants to see that something happens, so a semicolon will be printed for every script
                if (itemsPropEve.Value.Count == 2 && (obj is HmiScreen || obj is HmiScreenItemBase))
                {
                    tempListPropEve.Insert(0, itemsPropEve.Value[0]);
                    string script = itemsPropEve.Value[1];
                    tagSetUsage = tagSetUsage.Concat(GetTagSetUsage(script, objectName + delimiter + itemsPropEve.Key));
                    tempListPropEve.Insert(1, Environment.NewLine + "export function _" + objectName + "_" + itemsPropEve.Key + "_OnPropertyChanged() {" + script + Environment.NewLine + "}");
                }

                if (itemsPropEve.Value.Count == 1 && (obj is HmiScreen || obj is HmiScreenItemBase))
                {
                    tempListPropEve.Add(Environment.NewLine + "//eslint-disable-next-line camelcase");
                    string script = itemsPropEve.Value[0];
                    tagSetUsage = tagSetUsage.Concat(GetTagSetUsage(script, objectName + delimiter + itemsPropEve.Key));
                    tempListPropEve.Add("export function _" + objectName + "_" + itemsPropEve.Key + "_OnPropertyChanged() {" + script + Environment.NewLine + " }");
                }

                if (itemsPropEve.Value.Count == 2 && !(obj is HmiScreen || obj is HmiScreenItemBase))
                {
                    tempListPropEve.Insert(0, itemsPropEve.Value[0]);
                    string script = itemsPropEve.Value[1];
                    tagSetUsage = tagSetUsage.Concat(GetTagSetUsage(script, itemsPropEve.Key));
                    tempListPropEve.Insert(1, "_" + itemsPropEve.Key + "_OnPropertyChanged() {" + script + Environment.NewLine + "}");
                }

                if (itemsPropEve.Value.Count == 1 && !(obj is HmiScreen || obj is HmiScreenItemBase))
                {
                    tempListPropEve.Add(Environment.NewLine + "//eslint-disable-next-line camelcase");
                    string script = itemsPropEve.Value[0];
                    tagSetUsage = tagSetUsage.Concat(GetTagSetUsage(script, itemsPropEve.Key));
                    tempListPropEve.Add("_" + itemsPropEve.Key + "_OnPropertyChanged() {" + script + Environment.NewLine + "}");
                }
            }
            // Nodes for going recursive
            if (deepSearch)
            {
                var objKeys = from attributeInfo in obj.GetAttributeInfos()
                              where ((obj.GetAttribute(attributeInfo.Name) as IEngineeringObject) != null)
                              select attributeInfo.Name;
                var objProps = obj.GetAttributes(objKeys);


                foreach (IEngineeringObject item in objProps)
                {
                    if (item.GetType().Name != "MultilingualText")
                    {
                        var nodeDynPropEve = GetAllMyAttributesDynPropEves(item, deepSearch, whereConditions, sets, ref tagSetUsage, ref cycles);
                        foreach (var dyn in nodeDynPropEve[0])
                        {
                            Console.Write(';');  // the user wants to see that something happens, so a semicolon will be printed for every script
                            tempListDyn.Add("function _" + objectName + dyn);
                        }
                        int index = 0;
                        foreach (var propEve in nodeDynPropEve[1])
                        {
                            if (index != 0)
                            {
                                Console.Write(';');  // the user wants to see that something happens, so a semicolon will be printed for every script
                                tempListPropEve.Add(Environment.NewLine + "export function _" + obj.GetAttribute("Name") + propEve);
                            }
                            else
                            {
                                tagSetUsage = tagSetUsage.Concat(GetTagSetUsage(propEve, obj.GetAttribute("Name").ToString()));
                                tempListPropEve.Add(propEve);
                                index++;
                            }
                        }
                    }
                }
            }
            return new List<List<string>>() { tempListDyn, tempListPropEve };
        }

        private List<string> GetTagSetUsage(string script, string eventDynName)
        {
            var tagSetUsage = new List<string>();
            var scriptLines = script.Split('\n').Where(x => !string.IsNullOrWhiteSpace(x.Trim()));
            var keyWordsToTest = new Dictionary<string, int> { { ".Read(", 0 }, { ".Write(", 0 }, { ".SetTagValue(", 0 } };
            var keys = keyWordsToTest.Keys.ToList();
            foreach (var scriptLine in scriptLines)
            {
                foreach (var key in keys)
                {
                    if (keyWordsToTest[key] < 2 && scriptLine.Contains(key)) {
                        keyWordsToTest[key]++;
                        if (keyWordsToTest[key] >= 2)
                        {
                            tagSetUsage.Add(eventDynName);
                        }
                    }
                }
            }
            return tagSetUsage;
        }

        public List<string> GetAllMyAttributesEve(IEngineeringObject obj, List<string> whereConditions, List<string> sets, ref IEnumerable<string> tagSetUsage, ref List<string> cycles)
        {
            List<string> tempListEve = new List<string>();

            var events = GetEvents(obj, whereConditions, sets, ref cycles);
            var objectName = obj.GetAttribute("Name").ToString().Replace(' ', '_').Replace('-', '_').Replace('&', 'ß');
            foreach (var itemsEve in events)
            {
                if (itemsEve.Value.Count == 2)
                {
                    Console.Write(';');  // the user wants to see that something happens, so a semicolon will be printed for every script
                    //tempListEve.Add(Environment.NewLine + "//eslint-disable-next-line camelcase");
                    tempListEve.Insert(0, itemsEve.Value[0]);
                    string script = itemsEve.Value[1];
                    tagSetUsage = tagSetUsage.Concat(GetTagSetUsage(script, objectName + delimiter + itemsEve.Key));
                    tempListEve.Insert(1, Environment.NewLine + "export async function _" + objectName + "_" + itemsEve.Key + "() {" + script + Environment.NewLine + "}");
                }

                if (itemsEve.Value.Count == 1)
                {
                    Console.Write(';');  // the user wants to see that something happens, so a semicolon will be printed for every script
                    tempListEve.Add(Environment.NewLine + "//eslint-disable-next-line camelcase");
                    string script = itemsEve.Value[0];
                    tagSetUsage = tagSetUsage.Concat(GetTagSetUsage(script, objectName + delimiter + itemsEve.Key));
                    tempListEve.Add("export async function _" + objectName + "_" + itemsEve.Key + "() {" + script + Environment.NewLine + "}");
                }
            }
            return tempListEve;
        }


        void SetMyAttributesSimpleTypes(string keyToSet, string valueToSet, IEngineeringObject obj)
        {
            Type _type = obj.GetType().GetProperty(keyToSet)?.PropertyType;

            object attrVal = null;
            if (_type != null && _type.BaseType == typeof(Enum))
            {
                attrVal = Enum.Parse(_type, valueToSet);
            }
            else if (_type != null && _type.Name == "Color")
            {
                var hexColor = new ColorConverter();
                attrVal = (Color)hexColor.ConvertFromString(valueToSet.ToUpper());
            }
            else if (keyToSet == "InitialAddress")
            {
                attrVal = valueToSet.Substring(0, valueToSet.Length - 1);
            }
            else if (obj.GetType().Name == "MultilingualText")
            {
                obj = (obj as MultilingualText).Items.FirstOrDefault(item => item.Language.Culture.Name == keyToSet);

                if (obj == null)
                {
                    Console.WriteLine("Language " + keyToSet + " does not exist in this Runtime!");
                    return;
                }
                keyToSet = "Text";
                attrVal = valueToSet;
            }
            else if (keyToSet == "Tags")
            {
                attrVal = valueToSet.Split(',').ToList();
            }
            else
            {
                if (_type != null) attrVal = Convert.ChangeType(valueToSet, _type);
            }

            try
            {
                obj.SetAttribute(keyToSet.ToString(), attrVal);
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        private Dictionary<string, List<string>> GetScripts(IEngineeringObject obj, string compositionName, List<string> whereConditions, List<string> sets, ref List<string> cycles)
        {
            var dict = new Dictionary<string, List<string>>();
            var scripts = new List<string>();

            if (compositionName == "Dynamizations")
            {
                try
                {
                    foreach (var dyn in (obj as UIBase).Dynamizations) //Dynamizations
                    {
                        var listDyn = new List<string>();
                        if (dyn.DynamizationType.ToString() == "Script")
                        {
                            foreach (var whereCondition in whereConditions.Where(wc => wc.StartsWith("Dynamization.")))
                            {
                                object triggerProp = dyn.GetType().GetProperty("Trigger")?.GetValue(dyn);
                                if (triggerProp != null)
                                {
                                    string[] triggerPropertyValue = whereCondition.Split('.').Last().Split('=');
                                    string propValue = triggerProp.GetType().GetProperty(triggerPropertyValue[0]).GetValue(triggerProp).ToString();
                                    if (propValue == triggerPropertyValue[1])
                                    {
                                        foreach (var set in sets.Where(s => s.StartsWith("Dynamization.")))
                                        {
                                            string key = set.Split('=')[0].Split('.').Last();
                                            //   'HMI_${Name.split("_")[-2]}_${Name.split("_")[-1]}'
                                            string valueToSet = set.Split('=')[1];
                                            if (valueToSet.StartsWith("'") && valueToSet.EndsWith("'"))
                                            {
                                                valueToSet = valueToSet.Replace("\'", "");
                                                int startIndex;
                                                while ((startIndex = valueToSet.IndexOf("${")) >= 0)
                                                {
                                                    int endIndex = valueToSet.IndexOf("}", startIndex);
                                                    string subString = valueToSet.Substring(startIndex, endIndex - startIndex + 1);
                                                    string tempSubString = subString.Replace("${", "").Replace("}", "");
                                                    List<string> subStringSplit = tempSubString.Split('.').ToList();
                                                    string tempKey = subStringSplit[0];
                                                    subStringSplit.RemoveAt(0);
                                                    string tempValue = obj.GetAttribute(tempKey).ToString();
                                                    foreach (string part in subStringSplit)
                                                    {
                                                        if (part.StartsWith("split("))
                                                        {
                                                            int startIndex_ = part.IndexOf('(');
                                                            int endIndex_ = part.IndexOf(')', startIndex_);
                                                            char separator = part.Substring(startIndex_ + 1, endIndex_ - startIndex_)[0];
                                                            string[] splitValue = tempValue.Split(separator);
                                                            startIndex_ = part.IndexOf('[');
                                                            endIndex_ = part.IndexOf(']', startIndex_);
                                                            int getIndex = Convert.ToInt32(part.Substring(startIndex_ + 1, endIndex_ - startIndex_ - 1));
                                                            if (getIndex < 0)
                                                            {
                                                                getIndex = splitValue.Length + getIndex;
                                                                if (getIndex < 0)
                                                                {
                                                                    getIndex = 0;
                                                                }
                                                            }
                                                            tempValue = splitValue[getIndex];
                                                        }
                                                    }
                                                    valueToSet = valueToSet.Replace(subString, tempValue);
                                                }
                                            }
                                            SetMyAttributesSimpleTypes(key, valueToSet, triggerProp as IEngineeringObject);
                                        }
                                    }
                                }
                            }
                            if (string.IsNullOrEmpty(globalDefinitionAreaScriptCodeDyns))
                            {
                                globalDefinitionAreaScriptCodeDyns = dyn.GetType().GetProperty("GlobalDefinitionAreaScriptCode").GetValue(dyn).ToString().Replace("\r\n", Environment.NewLine);
                            }
                            string script = dyn.GetType().GetProperty("ScriptCode").GetValue(dyn).ToString().Replace("\r\n", Environment.NewLine);
                            if ((dyn as ScriptDynamization).Trigger.Type == TriggerType.Tags)
                            {
                                listDyn.Add("// Trigger: " + (dyn as ScriptDynamization).Trigger.Type.ToString() + " " + String.Join(", ", ((dyn as ScriptDynamization).Trigger.Tags) as List<string>) + Environment.NewLine + script);
                            }
                            else
                            {
                                listDyn.Add("// Trigger: " + (dyn as ScriptDynamization).Trigger.Type.ToString() + Environment.NewLine + script);
                                if ((dyn as ScriptDynamization).Trigger.Type != TriggerType.Disabled)
                                {
                                    cycles.Add(dyn.Parent.GetAttribute("Name") + delimiter + dyn.PropertyName + delimiter + (dyn as ScriptDynamization).Trigger.Type);
                                }
                            }
                            switch ((dyn as ScriptDynamization).Trigger.Type)
                            {
                                case TriggerType.CustomCycle:
                                    countDynTriggers["OtherCycles"]++;
                                    break;
                                case TriggerType.T100ms:
                                    countDynTriggers["T100ms"]++;
                                    break;
                                case TriggerType.T250ms:
                                    countDynTriggers["T250ms"]++;
                                    break;
                                case TriggerType.T500ms:
                                    countDynTriggers["T500ms"]++;
                                    break;
                                case TriggerType.T1s:
                                    countDynTriggers["T1s"]++;
                                    break;
                                case TriggerType.T2s:
                                    countDynTriggers["T2s"]++;
                                    break;
                                case TriggerType.T5s:
                                    countDynTriggers["T5s"]++;
                                    break;
                                case TriggerType.T10s:
                                    countDynTriggers["T10s"]++;
                                    break;
                                case TriggerType.Tags:
                                    countDynTriggers["Tags"]++;
                                    foreach(var tag in (dyn as ScriptDynamization).Trigger.Tags as List<string>)
                                    {
                                        if (!tagNames.Contains(tag))
                                        {
                                            tagNames.Add(tag);
                                        }
                                    }
                                    break;
                                case TriggerType.Disabled:
                                    countDynTriggers["Disabled"]++;
                                    break;
                            }
                        }
                        else if (dyn.DynamizationType.ToString() == "Tag")
                        {
                            if (!tagNames.Contains(dyn.GetAttribute("Tag").ToString()))
                            {
                                tagNames.Add(dyn.GetAttribute("Tag").ToString());
                            }
                            countDynTriggers["TagDynamizations"]++;
                        }
                        else if (dyn.DynamizationType.ToString() == "ResourceList")
                        {
                            countDynTriggers["ResourceLists"]++;
                        }
                        else
                        {
                            Console.WriteLine("Unknown dynamization type: " + dyn.DynamizationType.ToString());
                        }
                        dict.Add(dyn.PropertyName, listDyn);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            else // e.g. EventHandlers & PropertyEventHandlers
            {
                int i = 1;
                foreach (IEngineeringObject eveHand in obj.GetComposition(compositionName) as IEngineeringComposition)
                {
                    var listEve = new List<string>();
                    IEngineeringObject script = eveHand.GetAttribute("Script") as IEngineeringObject;
                    //script.SetAttribute("ScriptCode", "\r\n\nEureka");

                    if (string.IsNullOrEmpty(globalDefinitionAreaScriptCodeEvents))
                    {
                        string globalDef = script.GetAttribute("GlobalDefinitionAreaScriptCode").ToString().Replace("\r\n", Environment.NewLine);
                        if (!globalDef.Contains("[FunctionListModule]")) // there is no global definition area of a function list module...
                        {
                            globalDefinitionAreaScriptCodeEvents = globalDef;
                        }
                    }
                    listEve.Add(script.GetAttribute("ScriptCode").ToString().Replace("\r\n", Environment.NewLine));

                    if (compositionName == "EventHandlers")
                    {
                        if (dict.ContainsKey(eveHand.GetAttribute("EventType").ToString()))
                        {
                            dict.Add(eveHand.GetAttribute("EventType").ToString() + "_" + i, listEve);
                            i++;
                        }
                        else {
                            dict.Add(eveHand.GetAttribute("EventType").ToString(), listEve);
                        }             
                    }
                    else
                    {
                        dict.Add(eveHand.GetAttribute("PropertyName").ToString(), listEve);
                    }
                }
            }
            return dict;
        }
        private Dictionary<string, List<string>> GetDynamizations(IEngineeringObject obj, List<string> whereConditions, List<string> sets, ref List<string> cycles)
        {
            return GetScripts(obj, "Dynamizations", whereConditions, sets, ref cycles);
        }
        private Dictionary<string, List<string>> GetEvents(IEngineeringObject obj, List<string> whereConditions, List<string> sets, ref List<string> cycles)
        {
            return GetScripts(obj, "EventHandlers", whereConditions, sets, ref cycles);
        }
        private Dictionary<string, List<string>> GetPropertyEvents(IEngineeringObject obj, List<string> whereConditions, List<string> sets, ref List<string> cycles)
        {
            return GetScripts(obj, "PropertyEventHandlers", whereConditions, sets, ref cycles);
        }
    }

    class ScreenDynEvents
    {
        private string name;
        private List<string> dynList;
        private List<string> eventList;

        public ScreenDynEvents(string name, ref List<string> dynList, ref List<string> eventList)
        {
            this.name = name;
            this.dynList = dynList;
            this.eventList = eventList;
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public List<string> DynList
        {
            get { return dynList; }
            set { dynList = value; }
        }

        public List<string> EventList
        {
            get { return eventList; }
            set { eventList = value; }
        }
    }

}

using Siemens.Engineering;
using Siemens.Engineering.HmiUnified.UI;
using Siemens.Engineering.HmiUnified.UI.Base;
using Siemens.Engineering.HmiUnified.UI.Dynamization.Script;
using Siemens.Engineering.HmiUnified.UI.Screens;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using Siemens.Engineering.HmiUnified.UI.Controls;
using System.Globalization;
using System.Diagnostics;
using Siemens.Engineering.HmiUnified.UI.Dynamization;
using Siemens.Engineering.HmiUnified.UI.Events;
using System.IO;

namespace ShowScripts
{
    class AddScriptsToList
    {
        static List<string> screenItemEvents = new List<string>();
        static List<string> screenItemDynamisations = new List<string>();
        static List<string> tagNames = new List<string>();
        static Dictionary<string, int> countDynTriggers;
        static Dictionary<string, int> countScreenItems;
        static Dictionary<string, int> faceplateTypes;
        static string globalDefinitionAreaScriptCodeDyns;
        static string globalDefinitionAreaScriptCodeEvents;
        static List<string> childScreens;
        static int check = 0;


        public static void ImportScripts(IEnumerable<HmiScreen> screens, string fileDirectory)
        {
            int getLines = 0;
            DirectoryInfo d = new DirectoryInfo(fileDirectory);
            FileInfo[] Files = d.GetFiles();
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
                            IEngineeringObject script = eveHandler.GetAttribute("Script") as IEngineeringObject;
                            string eveType = eveHandler.GetAttribute("EventType").ToString();
                            List<string> scriptLines = new List<string>();
                            foreach (string line in lines)
                            {
                                if (line.Contains("function _" + screen.Name + "_" + eveType))
                                {
                                    getLines = 1;
                                }
                                if (getLines == 1)
                                {
                                    scriptLines.Add(line);
                                }
                                if (line.Contains("//eslint-disable-next-line camelcase") && getLines == 1)
                                {
                                    getLines = 0;
                                    scriptLines.Remove("");
                                    scriptLines.Remove("//eslint-disable-next-line camelcase");
                                    break;
                                }
                            }
                            getLines = 0;
                            scriptLines.RemoveAt(0); // remove the first line, because it is the header line
                            scriptLines.RemoveAt(scriptLines.Count-1); // remove the last line, because it is the closing bracket }
                            string setScript = string.Join(Environment.NewLine, scriptLines);
                            script.SetAttribute("ScriptCode", setScript);
                        }
                        foreach (PropertyEventHandler evePropHandler in screen.PropertyEventHandlers)
                        {
                            IEngineeringObject script = evePropHandler.GetAttribute("Script") as IEngineeringObject;
                            string eveType = evePropHandler.GetAttribute("EventType").ToString();
                            List<string> scriptLines = new List<string>();
                            foreach (string line in lines)
                            {
                                if (line.Contains(eveType))
                                {
                                    getLines = 1;
                                }
                                if (getLines == 1)
                                {
                                    scriptLines.Add(line);
                                }
                                if (line.Contains("//eslint-disable-next-line camelcase") && getLines == 1)
                                {
                                    getLines = 0;
                                    scriptLines.Remove("");
                                    scriptLines.Remove("//eslint-disable-next-line camelcase");
                                    break;
                                }
                            }
                            getLines = 0;
                            string setScript = "";
                            for (int i = 1; i < scriptLines.Count - 1; i++)
                            {
                                setScript = setScript + "\n" + scriptLines[i];
                            }
                            script.SetAttribute("ScriptCode", setScript);
                        }
                        HmiScreenItemBaseComposition items = screen.ScreenItems;
                        foreach (HmiScreenItemBase item in items)
                        {
                            foreach (PropertyEventHandler evePropHandlerItem in item.PropertyEventHandlers)
                            {
                                IEngineeringObject script = evePropHandlerItem.GetAttribute("Script") as IEngineeringObject;
                                string eveType = evePropHandlerItem.GetAttribute("EventType").ToString();
                                List<string> scriptLines = new List<string>();
                                foreach (string line in lines)
                                {
                                    if (line.Contains(eveType))
                                    {
                                        getLines = 1;
                                    }
                                    if (getLines == 1)
                                    {
                                        scriptLines.Add(line);
                                    }
                                    if (line.Contains("//eslint-disable-next-line camelcase") && getLines == 1)
                                    {
                                        getLines = 0;
                                        scriptLines.Remove("");
                                        scriptLines.Remove("//eslint-disable-next-line camelcase");
                                        break;
                                    }
                                }
                                getLines = 0;
                                string setScript = "";
                                for (int i = 1; i < scriptLines.Count - 1; i++)
                                {
                                    setScript = setScript + "\n" + scriptLines[i];
                                }
                                script.SetAttribute("ScriptCode", setScript);
                            }
                            Console.WriteLine(item.GetType());
                            IEngineeringObject itemConcreto = item;
                            Console.WriteLine(itemConcreto.GetAttribute("Name"));
                            Console.WriteLine(itemConcreto.GetComposition("EventHandlers"));
                            foreach (IEngineeringObject eveHandItem in itemConcreto.GetComposition("EventHandlers") as IEngineeringComposition)
                            {
                                Console.WriteLine(eveHandItem);
                                IEngineeringObject script = eveHandItem.GetAttribute("Script") as IEngineeringObject;
                                string eveType = eveHandItem.GetAttribute("EventType").ToString();
                                string itemName = itemConcreto.GetAttribute("Name").ToString().Replace(" ", "_");
                                List<string> scriptLines = new List<string>();
                                foreach (string line in lines)
                                {
                                    if (line.Contains(eveType) && line.Contains(itemName))
                                    {
                                        getLines = 1;
                                    }
                                    if (getLines == 1)
                                    {
                                        scriptLines.Add(line);
                                    }
                                    if (line.Contains("//eslint-disable-next-line camelcase") && getLines == 1)
                                    {
                                        getLines = 0;
                                        scriptLines.Remove("");
                                        scriptLines.Remove("//eslint-disable-next-line camelcase");
                                        break;
                                    }
                                }
                                getLines = 0;
                                string setScript = "";
                                for (int i = 1; i < scriptLines.Count - 1; i++)
                                {
                                    setScript = setScript + "\n" + scriptLines[i];
                                }
                                script.SetAttribute("ScriptCode", setScript);
                            }
                        }

                    }
                    else if (file.Name.EndsWith("_Dynamizations.js") && file.Name.StartsWith(screen.Name))
                    {
                        string tempPath = fileDirectory + file.Name;
                        string[] lines = System.IO.File.ReadAllLines(tempPath);
                        foreach (DynamizationBase dynamizationScreen in screen.Dynamizations)
                        {
                            Console.WriteLine(dynamizationScreen);
                            string propertyType = dynamizationScreen.GetAttribute("PropertyName").ToString();
                            //string script = dynamizationScreen.GetAttribute("ScriptCode").ToString();
                            IEngineeringObject script = dynamizationScreen.GetAttribute("ScriptCode") as IEngineeringObject;
                            List<string> scriptLines = new List<string>();
                            foreach (string line in lines)
                            {
                                if (line.Contains(propertyType))
                                {
                                    getLines = 1;
                                }
                                if (getLines == 1)
                                {
                                    scriptLines.Add(line);
                                }
                                if (line.Contains("//eslint-disable-next-line camelcase") && getLines == 1)
                                {
                                    getLines = 0;
                                    scriptLines.Remove("");
                                    scriptLines.Remove("//eslint-disable-next-line camelcase");
                                    break;
                                }
                            }
                            getLines = 0;
                            string setScript = "";
                            for (int i = 1; i < scriptLines.Count - 1; i++)
                            {
                                setScript = setScript + "\n" + scriptLines[i];
                            }
                            dynamizationScreen.SetAttribute("ScriptCode", setScript);
                            //script.SetAttribute("ScriptCode", setScript);
                        }
                        HmiScreenItemBaseComposition items = screen.ScreenItems;
                        foreach (HmiScreenItemBase item in items)
                        {
                            foreach (DynamizationBase dynamizationItem in item.Dynamizations)
                            {
                                Console.WriteLine(dynamizationItem);
                                if (dynamizationItem.DynamizationType == DynamizationType.Script)
                                {
                                    string propertyType = dynamizationItem.GetAttribute("PropertyName").ToString();
                                    List<string> scriptLines = new List<string>();
                                    foreach (string line in lines)
                                    {
                                        if (line.Contains(propertyType))
                                        {
                                            getLines = 1;
                                        }
                                        if (getLines == 1)
                                        {
                                            scriptLines.Add(line);
                                        }
                                        if (line.Contains("//eslint-disable-next-line camelcase") && getLines == 1)
                                        {
                                            getLines = 0;
                                            scriptLines.Remove("");
                                            scriptLines.Remove("//eslint-disable-next-line camelcase");
                                            break;
                                        }
                                    }
                                    getLines = 0;
                                    string setScript = "";
                                    for (int i = 1; i < scriptLines.Count - 1; i++)
                                    {
                                        setScript = setScript + "\n" + scriptLines[i];
                                    }
                                    dynamizationItem.SetAttribute("ScriptCode", setScript);
                                }
                            }
                        }
                    }
                }
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
        public static void ExportScripts(IEnumerable<HmiScreen> screens, string fileDirectory, string screenName, string whereCondition = "", string sets = "", bool deepSearch = false)
        {
            var csvStringP = new List<string>();
            string _screenName = "";
            
            if(screenName == ".*" && check == 0)
            {             
                //_screenName = Interaction.InputBox("Enter a screen name: ", "Screen name", "All screens");
                _screenName = Interaction.InputBox("Enter a screen name: ", "Screen name");
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

            csvStringP.Add("Screen name,Item count,Cyclic (<=500ms),Cyclic (>500ms),OtherCycles,Disabled,Tags-Dyn-Scripts,Tags-Dyn, Total Number of Tags,Resource list,Events,Child screens");

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
                csvStringP[0] += "," + entry.Key;
            }

            faceplateTypes = new Dictionary<string, int>();           

            foreach (var screen in screens.Where(s => Regex.Matches(s.Name, _screenName, RegexOptions.IgnoreCase).Count > 0))
            {
                var dynamizationList = new List<string>();
                var eveList = new List<string>();
                var screenDynEventList = new List<ScreenDynEvents>();
                // inits
                Console.WriteLine(screen.Name);
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

                foreach (var key in countScreenItems.Keys.ToList())
                {
                    countScreenItems[key] = 0;
                }

                foreach (var key in faceplateTypes.Keys.ToList())
                {
                    faceplateTypes[key] = 0;
                }

                // calculations
                var screenDynsPropEves = GetAllMyAttributesDynPropEves(screen, deepSearch, whereCondition.Split(',').ToList(), sets.Split(',').ToList());
                var screenDyns = screenDynsPropEves[0];
                var screenPropEves = screenDynsPropEves[1];
                var screenEves = GetAllMyAttributesEve(screen, whereCondition.Split(',').ToList(), sets.Split(',').ToList());

                foreach (var screenitem in screen.ScreenItems)
                {
                    var screenitemDynsPropEves = GetAllMyAttributesDynPropEves(screenitem, deepSearch, whereCondition.Split(',').ToList(), sets.Split(',').ToList());
                    var screenitemDyns = screenitemDynsPropEves[0];
                    var screenitemPropEves = screenitemDynsPropEves[1];
                    var screenitemEves = GetAllMyAttributesEve(screenitem, whereCondition.Split(',').ToList(), sets.Split(',').ToList());

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
                            csvStringP[0] += "," + faceplate.ContainedType;
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
                csvStringP.Add(screen.Name + "," + screen.ScreenItems.Count + "," + (countDynTriggers["T100ms"] + countDynTriggers["T250ms"] + countDynTriggers["T500ms"]) + "," +
                    (countDynTriggers["T1s"] + countDynTriggers["T2s"] + countDynTriggers["T5s"] + countDynTriggers["T10s"]) + "," + countDynTriggers["OtherCycles"] + "," +
                    countDynTriggers["Disabled"] + "," + countDynTriggers["Tags"] + "," + countDynTriggers["TagDynamizations"] + "," + tagNames.Count() + "," + countDynTriggers["ResourceLists"] + "," +
                    eventListCount + "," + string.Join("&", childScreens) + ",");

                foreach (var entry in countScreenItems)
                {
                    csvStringP[csvStringP.Count - 1] += entry.Value + ",";
                }

                foreach (var entry in faceplateTypes)
                {
                    csvStringP[csvStringP.Count - 1] += entry.Value + ",";
                }

                using (StreamWriter sw = new StreamWriter(fileDirectory + screen.Name + "_Dynamizations.js"))
                {
                    sw.Write(string.Join(Environment.NewLine, dynList));
                }
                using (StreamWriter sw = new StreamWriter(fileDirectory + screen.Name + "_Events.js"))
                {
                    sw.Write(string.Join(Environment.NewLine, eveList));
                }
            }

            //fill empty space with 0 
            int longestEntry = csvStringP[0].Count(x => x == ',');
            int counterFaceplateTypes = 0;
            foreach (var entry in csvStringP.ToList())
            {
                if (counterFaceplateTypes != 0)
                {
                    for (int i = entry.Count(x => x == ','); i <= longestEntry; i++)
                    {
                        csvStringP[counterFaceplateTypes] += "0,";
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
        }

        static public List<List<string>> GetAllMyAttributesDynPropEves(IEngineeringObject obj, bool deepSearch, List<string> whereConditions, List<string> sets)
        {
            var tempListDyn = new List<string>();
            var tempListPropEve = new List<string>();

            var dynamizations = GetDynamizations(obj, whereConditions, sets);
            var propertyEvents = GetPropertyEvents(obj, whereConditions, sets);

            var objectName = obj.GetAttribute("Name").ToString().Replace(' ', '_').Replace('-', '_').Replace('&', 'ß');
            foreach (var itemsDyn in dynamizations)
            {
                if (itemsDyn.Value.Count == 2 && (obj is HmiScreen || obj is HmiScreenItemBase))
                {
                    Console.WriteLine("4");
                    tempListDyn.Insert(0, itemsDyn.Value[0]);
                    tempListDyn.Insert(1, "function _" + objectName + "_" + itemsDyn.Key + "_Trigger() {" + itemsDyn.Value[1] + Environment.NewLine + "}");
                }

                if (itemsDyn.Value.Count == 1 && (obj is HmiScreen || obj is HmiScreenItemBase))
                {
                    tempListDyn.Add(Environment.NewLine + "//eslint-disable-next-line camelcase");
                    tempListDyn.Add("function _" + objectName + "_" + itemsDyn.Key + "_Trigger() {" + itemsDyn.Value[0] + Environment.NewLine + "}");
                }

                if (itemsDyn.Value.Count == 1 && !(obj is HmiScreen || obj is HmiScreenItemBase))
                {
                    tempListDyn.Add(Environment.NewLine + "//eslint-disable-next-line camelcase");
                    tempListDyn.Add("_" + itemsDyn.Key + "_Trigger() {" + Environment.NewLine + itemsDyn.Value[0] + Environment.NewLine + "}");
                }
            }

            foreach (var itemsPropEve in propertyEvents)
            {
                if (itemsPropEve.Value.Count == 2 && (obj is HmiScreen || obj is HmiScreenItemBase))
                {
                    Console.WriteLine("3");
                    tempListPropEve.Insert(0, itemsPropEve.Value[0]);
                    tempListPropEve.Insert(1, Environment.NewLine + "export function _" + objectName + "_" + itemsPropEve.Key + "_OnPropertyChanged() {" + itemsPropEve.Value[1] + Environment.NewLine + "}");
                }

                if (itemsPropEve.Value.Count == 1 && (obj is HmiScreen || obj is HmiScreenItemBase))
                {
                    tempListPropEve.Add(Environment.NewLine + "//eslint-disable-next-line camelcase");
                    tempListPropEve.Add("export function _" + objectName + "_" + itemsPropEve.Key + "_OnPropertyChanged() {" + itemsPropEve.Value[0] + Environment.NewLine + " }");
                }

                if (itemsPropEve.Value.Count == 2 && !(obj is HmiScreen || obj is HmiScreenItemBase))
                {
                    Console.WriteLine("2");
                    tempListPropEve.Insert(0, itemsPropEve.Value[0]);
                    tempListPropEve.Insert(1, "_" + itemsPropEve.Key + "_OnPropertyChanged() {" + itemsPropEve.Value[1] + Environment.NewLine + "}");
                }

                if (itemsPropEve.Value.Count == 1 && !(obj is HmiScreen || obj is HmiScreenItemBase))
                {
                    tempListPropEve.Add(Environment.NewLine + "//eslint-disable-next-line camelcase");
                    tempListPropEve.Add("_" + itemsPropEve.Key + "_OnPropertyChanged() {" + itemsPropEve.Value[0] + Environment.NewLine + "}");
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
                        var nodeDynPropEve = GetAllMyAttributesDynPropEves(item, deepSearch, whereConditions, sets);
                        foreach (var dyn in nodeDynPropEve[0])
                        {
                            tempListDyn.Add("function _" + objectName + dyn);
                        }
                        int index = 0;
                        foreach (var propEve in nodeDynPropEve[1])
                        {
                            if (index != 0)
                            {
                                tempListPropEve.Add(Environment.NewLine + "export function _" + obj.GetAttribute("Name") + propEve);
                            }
                            else
                            {
                                tempListPropEve.Add(propEve);
                                index++;
                            }
                        }
                    }
                }
            }
            return new List<List<string>>() { tempListDyn, tempListPropEve };
        }

        static public List<string> GetAllMyAttributesEve(IEngineeringObject obj, List<string> whereConditions, List<string> sets)
        {
            List<string> tempListEve = new List<string>();

            var events = GetEvents(obj, whereConditions, sets);
            var objectName = obj.GetAttribute("Name").ToString().Replace(' ', '_').Replace('-', '_').Replace('&', 'ß');
            foreach (var itemsEve in events)
            {
                if (itemsEve.Value.Count == 2)
                {
                    Console.WriteLine("1");
                    //tempListEve.Add(Environment.NewLine + "//eslint-disable-next-line camelcase");
                    tempListEve.Insert(0, itemsEve.Value[0]);
                    tempListEve.Insert(1, Environment.NewLine + "export async function _" + objectName + "_" + itemsEve.Key + "() {" + itemsEve.Value[1] + Environment.NewLine + "}");
                }

                if (itemsEve.Value.Count == 1)
                {
                    tempListEve.Add(Environment.NewLine + "//eslint-disable-next-line camelcase");
                    tempListEve.Add("export async function _" + objectName + "_" + itemsEve.Key + "() {" + itemsEve.Value[0] + Environment.NewLine + "}");
                }
            }
            return tempListEve;
        }


        static void SetMyAttributesSimpleTypes(string keyToSet, string valueToSet, IEngineeringObject obj)
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

        private static Dictionary<string, List<string>> GetScripts(IEngineeringObject obj, string compositionName, List<string> whereConditions, List<string> sets)
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
                            if ((dyn as ScriptDynamization).Trigger.Type.ToString() == "Tags")
                            {
                                listDyn.Add("// Trigger: " + (dyn as ScriptDynamization).Trigger.Type.ToString() + " " + String.Join(", ", ((dyn as ScriptDynamization).Trigger.Tags) as List<string>) + Environment.NewLine + script);
                            }
                            else
                            {
                                listDyn.Add("// Trigger: " + (dyn as ScriptDynamization).Trigger.Type.ToString() + Environment.NewLine + script);
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
                foreach (IEngineeringObject eveHand in obj.GetComposition(compositionName) as IEngineeringComposition)
                {
                    int i = 1;
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
        private static Dictionary<string, List<string>> GetDynamizations(IEngineeringObject obj, List<string> whereConditions, List<string> sets)
        {
            return GetScripts(obj, "Dynamizations", whereConditions, sets);
        }
        private static Dictionary<string, List<string>> GetEvents(IEngineeringObject obj, List<string> whereConditions, List<string> sets)
        {
            return GetScripts(obj, "EventHandlers", whereConditions, sets);
        }
        private static Dictionary<string, List<string>> GetPropertyEvents(IEngineeringObject obj, List<string> whereConditions, List<string> sets)
        {
            return GetScripts(obj, "PropertyEventHandlers", whereConditions, sets);
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

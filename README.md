# TIA Add-In ShowScripts

Export all JavaScripts of all screens to a file per screen in your project UserFiles' directory and get an additional Excel overview over all screens.

# Development

Open TestAddin.sln and develop your code in AddScriptsToList.cs

# Deployment

Open ShowScripts.sln, build as release and deliver ShowScripts.addin
Alternative, build a release of TestAddin.sln to get a console program to automate your workflow

# Installation and usage of TIA Portal Add-Ins
https://support.industry.siemens.com/cs/ww/de/view/109773999 \
Please refer to the above link and follow the instructions for installation.

# Run Add-In for WinCC Unified Devices
Select the device you would like to analyze. \
Right mouse click and navigate to the "Add-Ins" Chapter. \
Select "ShowScriptCode" -> "Show all Scripts of HMI" and execute it.

# Screen selection 
After the function is executed another popup window will appear. \
(If it is not coming up check the taskbar)\
In this window you can select to export information about a specific screen by entering the screen name.\
If no information is hand over all screens will be exported.

# Result 
The export does require some time. Depending on your project it can take up to 30 Minutes.\
During the export the TIA Portal cannot be used for further actions. \
The exported information can be found in the "UserFiles" of the project.\
E.g. "C:\User\siemens\Documents\Automation\ProjextXX\UserFiles\"

# TestAddin console options

| Option          | Beschreibung |
:---------------  | :------------
| \-ALLSCREENS    | all screens of your project will be exported, without that you can specify your screen in a following input prompt |
| \-C             | closes the TIA Portal project at the end |
| \-D             | specifies the export directory, if unused the UserFiles directory in the TIA project will be taken, example: \-D C:\\temp\\ |
| \-H 			  | calls the help, additionally possible by \-? |
| \-?			  | calls the help, additionally possible by \-H |
| \-P			  | specifies the TIA project to use, if unused the already opened project will be used, example: \-P D:\\TiaProjects\\Digi.ap17\\ |
| \-Q			  | closes the TIA Portal instance at the end |
| \-SET			  | \-SET Dynamization.Trigger.Type=4, Dynamization.Trigger.Tags='Refresh_tag' \-WHERE Dynamization.Trigger.Type=250 |
| \-U			  | only if no TIA Portal is startet, it starts TIA Portal with user interface |
| \-UPDATE		  | specifies a single screen, to only export that screen, example: \-U Screen_1 |
| \-WHERE		  | can be used as a filter |



# Limitations 
The tool is based on TIA Portal Openness and does have limitations.
- Scheduled tasks and global module content are **not** accessible
- ScreenItems: "My controls"/CustomWebControls, Symbolic IO field, Touch Area and DynamicSVG are **not** supported
- Text- und Graphiclists are **not** accessible
- Only Faceplate-Containers can be checked
- Library-Handling/MasterCopy-Handling (CopyFrom, CopyTo… whole Unified device only)
- Folder / Folder structure of Screens and Tag tables
- ScreenItem Properties: Layer
- Rename Unified RT (Rename PC RT does work)
- Create integrated connections (only not integrated connections are supported)





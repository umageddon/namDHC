#singleInstance off
#noEnv
#Persistent
detectHiddenWindows On
setTitleMatchmode 3
SetBatchLines, -1
SetKeyDelay, -1, -1
SetMouseDelay, -1
SetDefaultMouseSpeed, 0
SetWinDelay, -1
SetControlDelay, -1

/*
 v1.0		- Initial release

 v1.01		- Added ISO input media support
			- Minor fixes
			
 v1.02		- Removed superfluous code
			- Some spelling mistakes fixed
			- Added some comments
			- Minor GUI changes
 
*/

#Include SelectFolderEx.ahk
#Include ClassImageButton.ahk
#Include ConsoleClass.ahk
#Include JSON.ahk

onExit("quitApp", 1)

; Default global values 
; --------------
mainAppVersion := "1.01"
chdmanLocation := a_scriptDir "\chdman.exe"
chdmanVerArray := ["0.236", "0.237", "0.238", "0.239", "0.240"]
mainAppName := "namDHC"
mainAppNameVerbose := mainAppName " - Verbose"
runAppName := mainAppName " - Job"																																																			
runAppNameChdman := runAppName " - chdman"
runAppNameConsole := runAppName " - Console"
jobTimeoutSec := 3
chdmanTimeoutTimeSec := 5
waitTimeConsoleSec := 15
chdmanOptionMaxPerSide := 9
jobQueueSize := 3
jobQueueSizeLimit := 10
outputFolder := a_workingDir
playFinishedSong := "yes"
removeFileEntryAfterFinish := "yes"
showJobConsole := "no"
showVerboseWin := "no"
verboseWinPosH := 400 
verboseWinPosW := 800
verboseWinPosX := 775
verboseWinPosY := 150
mainWinPosX := 800
mainWinPosY := 100

ini("read", ["jobQueueSize"
			,"outputFolder"
			,"showJobConsole"
			,"showVerboseWin"
			,"playFinishedSong"
			,"removeFileEntryAfterFinish"
			,"mainWinPosX"
			,"mainWinPosY"
			,"verboseWinPosW"
			,"verboseWinPosH"
			,"verboseWinPosX"
			,"verboseWinPosY"])

if ( !fileExist(chdmanLocation) ) {
	msgbox 16, % "Fatal Error", % "CHDMAN.EXE not found!`n`nMake sure the chdman executable is located in the same directory as namDHC and try again.`n`nThe following chdman verions are supported:`n" arrayToComma(chdmanVerArray)
	exitApp
}

;-------------------------------------------------------------
; Run a chdman thread
;-------------------------------------------------------------
if ( a_args[1] == "threadMode" ) {
	#include threads.ahk
}

killAllProcess()

scannedFiles := {}, chdmanOpt := {}, queuedData := []
GUI := { gVars:{}, dropdowns:{job:{}, media:{}}, buttons:{normal:[], hover:[], clicked:[], disabled:[]}, menu:{namesOrder:[], File:[], Settings:[], About:[]} }
GUI.dropdowns["job"] := {create:"Create CHD files from media", extract:"Extract images from CHD files", info:"Get info from CHD files", verify:"Verify CHD files", addMeta:"Add metadata to CHD files", delMeta:"Delete metadata from CHD files"}
GUI.dropdowns["media"] := {cd:"CD image", hd:"Hard disk image", ld:"LaserDisc image", raw:"Raw image"}
GUI.buttons["default"] := {normal:[0, 0xFFCCCCCC, "", "", 3], 			hover:[0, 0xFFBBBBBB, "", 0xFF555555, 3], 	clicked:[0, 0xFFCFCFCF, "", 0xFFAAAAAA, 3], disabled:[0, 0xFFE0E0E0, "", 0xFFAAAAAA, 3] }
GUI.buttons["cancel"]  := {normal:[0, 0xFFFC6D62, "", "White", 3], 		hover:[0, 0xFFff8e85, "", "White", 3], 		clicked:[0, 0xFFfad5d2, "", "White", 3], 	disabled:[0, 0xFFfad5d2, "", "White", 3]}
GUI.buttons["start"]   := {normal:[0, 0xFF74b6cc, "", 0xFF444444, 3],	hover:[0, 0xFF84bed1, "", "White", 3], 		clicked:[0, 0xFFa5d6e6, "", "White", 3], 	disabled:[0, 0xFFd3dde0, "", 0xFF888888, 3] }	
GUI.menu["namesOrder"] := ["File", "Settings", "About"]
GUI.menu.File[1] := {name:"Quit", 													gotolabel:"quitApp", 					saveVar:""}
GUI.menu.About[1] := {name:"About", 												gotolabel:"menuSelected", 				saveVar:""}
GUI.menu.Settings[1] := {name:"Number of jobs to run concurrently", 				gotolabel:":SubSettingsConcurrently", 	saveVar:""}
GUI.menu.Settings[2] := {name:"Show verbose window", 								gotolabel:"menuSelected", 				saveVar:"showVerboseWin", Fn:"showVerboseWindow"}
GUI.menu.Settings[3] := {name:"Show a console for each job",						gotolabel:"menuSelected", 				saveVar:"showJobConsole"}
GUI.menu.Settings[4] := {name:"Play sounds when finished job queue", 				gotolabel:"menuSelected", 				saveVar:"playFinishedSong"}
GUI.menu.Settings[5] := {name:"Remove file entry from list when successful", 		gotolabel:"menuSelected", 				saveVar:"removeFileEntryAfterFinish"}
GUI.gVars.guiDefaultFont := guiDefaultFont()
GUI.gVars.templateHDDropdownList :=  ""											; Hard drive template dropdown list
. "|Conner CFA170A  -  163MB||"
. "Rodime R0201  -  5MB|"
. "Rodime R0202  -  10MB|"
. "Rodime R0203  -  15MB|"
. "Rodime R0204  -  20MB|"
. "Seagate ST-213  -  10MB|"
. "Segate ST-225  -  20MB|"
. "Seagate ST-251  -  40MB|"
. "Seagate ST-3600N  -  487MB|"
. "Maxtor LXT-213S  -  238MB|"
. "Maxtor LXT-340S  -  376MB|"
. "Maxtor MXT-540SL  -  733MB|"
. "Micropolis 1528  -  1272MB|"
chdmanOpt.force := 				{name: "force", 			paramString: "f", 	description: "Force overwriting an existing output file"}
chdmanOpt.verbose := 			{name: "verbose", 			paramString: "v", 	description: "Verbose output", 									hidden: true}
chdmanOpt.outputBin := 			{name: "outputbin", 		paramString: "ob", 	description: "Output filename for binary data", 				editField: "filename.bin", useQuotes:true}
chdmanOpt.inputParent := 		{name: "inputparent", 		paramString: "ip", 	description: "Input Parent", 									editField: "filename.ext", useQuotes:true}
chdmanOpt.inputStartFrame := 	{name: "inputstartframe", 	paramString: "isf", description: "Input Start Frame", 								editField: 0}
chdmanOpt.inputFrames := 		{name: "inputframes", 		paramString: "if", 	description: "Effective length of input in frames", 			editField: 0}
chdmanOpt.inputStartByte := 	{name: "inputstartbyte", 	paramString: "isb", description: "Starting byte offset within the input", 			editField: 0}
chdmanOpt.outputParent := 		{name: "outputparent",		paramString: "op", 	description: "Output parent file for CHD", 						editField: "filename.chd", useQuotes:true}
chdmanOpt.hunkSize := 			{name: "hunksize", 			paramString: "hs", 	description: "Size of each hunk (in bytes)", 					editField: 19584}
chdmanOpt.inputStartHunk := 	{name: "inputstarthunk",	paramString: "ish", description: "Starting hunk offset within the input", 			editField: 0}
chdmanOpt.inputBytes := 		{name: "inputBytes",		paramString: "ib", 	description: "Effective length of input (in bytes)", 			editField: 0}
chdmanOpt.compression := 		{name: "compression",		paramString: "c", 	description: "Compression codecs to use", 						editField: "cdlz,cdzl,cdfl"}
chdmanOpt.inputHunks := 		{name: "inputhunks",		paramString: "ih", 	description: "Effective length of input (in hunks)", 			editField: 0}
chdmanOpt.numProcessors := 		{name :"numprocessors",		paramString: "np", 	description: "Max number of CPU threads to use", 				dropdownOptions: procCountDDList()}
chdmanOpt.template := 			{name: "template", 			paramString: "tp",	description: "Hard drive template to use", 						dropdownOptions: GUI.gVars.templateHDDropDownList, dropdownValues:[0,1,2,3,4,5,6,7,8,9,10,11,12]}
chdmanOpt.chs := 				{name: "chs", 				paramString: "chs", description: "CHS Values [cyl, heads, sectors]", 				editField: "332,16,63"}
chdmanOpt.ident := 				{name: "ident", 			paramString: "id", 	description: "Name of ident file for CHS info", 				editField: "filename.chs", useQuotes:true}
chdmanOpt.size := 				{name: "size", 				paramString: "s", 	description: "Size of output file (in bytes)", 					editField: 0}
chdmanOpt.unitSize := 			{name: "unitsize", 			paramString: "us", 	description: "Size of each unit (in bytes)", 					editField: 0}
chdmanOpt.sectorSize := 		{name: "sectorsize", 		paramString: "ss", 	description: "Size of each hard disk sector (in bytes)", 		editField: 512}
chdmanOpt.deleteInputFiles := 	{name: "deleteInputFiles", 						description: "Delete input files after completing job", 		masterOf:"deleteInputDir"}
chdmanOpt.deleteInputDir := 	{name: "deleteInputDir", 						description: "Also delete input directory", 			xInset:10}
chdmanOpt.createSubDir :=		{name: "createSubDir",							description: "Create a new directory for each job"}
chdmanOpt.keepIncomplete :=		{name: "keepIncomplete",						description: "Keep failed or cancelled output files"}
;chdmanOpt.input := 			{name: "input", 			paramString: "i", 	description: "Input filename", 									hidden: true}
;chdmanOpt.output := 			{name: "output", 			paramString: "o", 	description: "Output filename", 								hidden: true}


createMainGUI()
createProgressBars() 
createMenus()
showVerboseWindow(showVerboseWin)											; Check or uncheck item "Show verbose window"  and show the window 

selectJob()																	; Select 1st selection in job dropdown list and trigger refreshGUI()

mainAppHWND := winExist(mainAppName)
mainAppMenuGet := DllCall("GetMenu", "uint", mainAppHWND)					; Save menu to retrieve later
mainMenuVisible := true

onMessage(0x03,		"moveGUIWin") 											; Windows are being moved - save positions in moveGUIWin()
onMessage(0x004A,	"receiveData")											; Recieving messages from threads

log(mainAppName " ready.")

return


; A Menu item was selected
;-------------------------
menuSelected() 
{
	global
	local selMenuObj, varName, fn
	
	switch a_ThisMenu {
		case "SettingsMenu":
			selMenuObj := GUI.menu.settings[a_ThisMenuItemPos]									; Reference menu setting
			varName := selMenuObj.saveVar														; Get variable name
			%varName% := (%varName% == "no")? "yes":"no"										; Toggle variable setting
			menu, SettingsMenu, % ((%varName% == "yes")? "Check":"UnCheck"), % selMenuObj.name	; Check or uncheck in menu
			ini("write", varName)																; Write new setting
			if ( isFunc(selMenuObj.Fn) ) {														; Check if function needs to be called
				fn := selMenuObj.Fn
				%fn%(%varName%)																	; Call function
			}

		case "SubSettingsConcurrently":											; Menu: Settings: User selected number of jobs to run concurrently
			loop % jobQueueSizeLimit											; Uncheck all 
				menu, SubSettingsConcurrently, UnCheck, % a_index
			menu, SubSettingsConcurrently, Check, % A_ThisMenuItemPos			; Check selected
			jobQueueSize := A_ThisMenuItemPos									; Set variable
			ini("write", "jobQueueSize")
			log("Saved jobQueueSize")
		
		case "AboutMenu":														; Menu: About
			gui 1:+OwnDialogs 
			msgbox, 64, % "About", % mainAppName " v" mainAppVersion "`nCopyright (C) 2021 Stuff 'n Things with Other Crap Inc."
	}
}



; User pressed input or output files button
; Show Ext menu
; --------------------------------------------
buttonExtSelect()
{
	switch a_guicontrol {
		case "buttonInputExtSelect":
			menu, InputExtTypes, Show, 873, 172 					; Hardcoded x,y positions as element position returns are wonky...
		case "buttonOutputExtSelect":
			menu, OutputExtTypes, Show, 873, 480
	}
}



; User selected an extension from the input/output extension menu
; ------------------------------------------------------------------
menuExtHandler()
{
	global
	local checkList, buildList, type
	
	checkList := "job" strReplace(a_ThisMenu, "ExtTypes") "Exts"

	if ( a_ThisMenu == "OutputExtTypes" ) {
		for idx, val in %checkList%
			menu, % a_ThisMenu, Uncheck, % val					; Uncheck all menu items, then check what user clicked ...
		menu, % a_ThisMenu, Check, % a_ThisMenuItem				;  ... so only one can be allowed to be selected
	}
	else if ( a_ThisMenu == "InputExtTypes" ) 
		menu, % a_ThisMenu, Togglecheck, % a_ThisMenuItem		; Let autohotkey take care of checking or unchecking input menu item
		
	type := strReplace(a_ThisMenu, "ExtTypes")
	buildList := "selected" type "Ext" 	
	%buildList% := []											; Build either of these Arrays: selectedOutputExt or selectedInputExt
	for idx, val in %checkList% {
		if ( isMenuChecked(a_ThisMenu, idx) ) {
			%buildList%.push(val)
		}
	}
	if ( %buildList%.length() == 0 ) {
		menu, % a_ThisMenu, check, % a_ThisMenuItem				; Make sure at least one item is checked
		%buildList%.push(a_ThisMenuItem)
	}
	
	guiCtrl({(buildList) "Text": arrayToComma(%buildList%)})
}


; Drop down job and media selections
; ------------------------------------
selectJob()
{
	global
	gui 1:submit, nohide

	switch dropdownMedia {
		case GUI.dropdowns.media.cd: 	jobMedia := "cd" 
		case GUI.dropdowns.media.hd:	jobMedia := "hd"
		case GUI.dropdowns.media.ld:	jobMedia := "ld"
		case GUI.dropdowns.media.raw:	jobMedia := "raw"
	}
	
	switch dropdownJob {																													
		case GUI.dropdowns.job.create:		jobCmd := "create" jobMedia, 	jobDesc := "Create CHD from a " stringUpper(jobMedia) " image", 	jobFinPreTxt := "Jobs created"			
		case GUI.dropdowns.job.extract:		jobCmd := "extract" jobMedia, 	jobDesc := "Extract a " stringUpper(jobMedia) " image from CHD",	jobFinPreTxt := "Jobs extracted"
		case GUI.dropdowns.job.info: 		jobCmd := "info", 				jobDesc := "Get info from CHD",										jobFinPreTxt := "Read info from jobs"
		case GUI.dropdowns.job.verify: 		jobCmd := "verify",				jobDesc := "Verify CHD",											jobFinPreTxt := "Jobs verified"
		case GUI.dropdowns.job.addMeta: 	jobCmd := "addmeta", 			jobDesc := "Add Metadata to CHD",									jobFinPreTxt := "Jobs with metadata added to"
		case GUI.dropdowns.job.delMeta: 	jobCmd := "delmeta", 			jobDesc := "Delete Metadata from CHD",								jobFinPreTxt := "Jobs with metadata deleted from"
	}
	
	switch jobCmd {
		case "extractcd": 	jobInputExts := ["chd"],				jobOutputExts := ["cue", "toc", "gdi"],	jobOptions := [chdmanOpt.force, chdmanOpt.createSubDir, chdmanOpt.deleteInputFiles, chdmanOpt.keepIncomplete, chdmanOpt.outputBin, chdmanOpt.inputParent]
		case "extractld": 	jobInputExts := ["chd"], 				jobOutputExts := ["raw"],				jobOptions := [chdmanOpt.force, chdmanOpt.deleteInputFiles, chdmanOpt.keepIncomplete, chdmanOpt.inputParent, chdmanOpt.inputStartFrame, chdmanOpt.inputFrames]
		case "extracthd": 	jobInputExts := ["chd"], 				jobOutputExts := ["img"],				jobOptions := [chdmanOpt.force, chdmanOpt.deleteInputFiles, chdmanOpt.keepIncomplete, chdmanOpt.inputParent, chdmanOpt.inputStartByte, chdmanOpt.inputStartHunk, chdmanOpt.inputBytes, chdmanOpt.inputHunks]
		case "extractraw":	jobInputExts := ["chd"], 				jobOutputExts := ["img", "raw"],		jobOptions := [chdmanOpt.force, chdmanOpt.deleteInputFiles, chdmanOpt.keepIncomplete, chdmanOpt.inputParent, chdmanOpt.inputStartByte, chdmanOpt.inputStartHunk, chdmanOpt.inputBytes, chdmanOpt.inputHunks]
		case "createcd": 	jobInputExts := ["cue", "toc", "gdi", "iso"],	jobOutputExts := ["chd"],			jobOptions := [chdmanOpt.force, chdmanOpt.deleteInputFiles, chdmanOpt.deleteInputDir, chdmanOpt.keepIncomplete, chdmanOpt.numProcessors, chdmanOpt.outputParent, chdmanOpt.hunkSize, chdmanOpt.compression]
		case "createld":	jobInputExts := ["raw"],				jobOutputExts := ["chd"],				jobOptions := [chdmanOpt.force, chdmanOpt.deleteInputFiles, chdmanOpt.deleteInputDir, chdmanOpt.keepIncomplete, chdmanOpt.numProcessors, chdmanOpt.outputParent, chdmanOpt.inputStartFrame, chdmanOpt.inputFrames, chdmanOpt.hunkSize, chdmanOpt.compression]
		case "createhd":	jobInputExts := ["img"],				jobOutputExts := ["chd"],				jobOptions := [chdmanOpt.force, chdmanOpt.deleteInputFiles, chdmanOpt.deleteInputDir, chdmanOpt.keepIncomplete, chdmanOpt.numProcessors, chdmanOpt.compression, chdmanOpt.outputParent, chdmanOpt.size, chdmanOpt.inputStartByte, chdmanOpt.inputStartHunk, chdmanOpt.inputBytes, chdmanOpt.inputHunks, chdmanOpt.hunkSize, chdmanOpt.ident, chdmanOpt.template, chdmanOpt.chs, chdmanOpt.sectorSize]
		case "createraw":	jobInputExts := ["img", "raw"],			jobOutputExts := ["chd"],				jobOptions := [chdmanOpt.force, chdmanOpt.deleteInputFiles, chdmanOpt.deleteInputDir, chdmanOpt.keepIncomplete, chdmanOpt.numProcessors, chdmanOpt.outputParent, chdmanOpt.inputStartByte, chdmanOpt.inputStartHunk, chdmanOpt.inputBytes, chdmanOpt.inputHunks, chdmanOpt.hunkSize, chdmanOpt.unitSize, chdmanOpt.compression]
		case "info":		jobInputExts := ["chd"],				jobOutputExts := [""],					jobOptions := []
		case "verify":		jobInputExts := ["chd"],				jobOutputExts := [""],					jobOptions := []
	}
	
	refreshGUI(true)
}

	


; Scan files and add to queue
; ----------------------------------
addFolderFiles()
{
	global
	local newFiles := [], extList := "", numAdded := 0, ext, inputFolder, path, idx, newInputList, thisFile
	
	gui 1:submit, nohide
	gui 1:+OwnDialogs 
	
	guiToggle("disable", "all")
	
	if ( !isObject(scannedFiles[jobCmd]) )
		scannedFiles[jobCmd] := []
	
	switch ( a_GuiControl ) {
		case "buttonAddFiles":
			for idx, ext in selectedInputExt
				extList .= "*." ext ";"
			fileSelectFile, newInputList, M3, % "::{20d04fe0-3aea-1069-a2d8-08002b30309d}", % "Select files", % extList
			if ( !errorLevel )
				loop, parse, newInputList, % "`n" 
				{
					if ( a_index == 1 )
						path := regExReplace(a_Loopfield, "\\$")
					else 
						newFiles.push(path "\" a_Loopfield)
				}
		
		case "buttonAddFolder":
			inputFolder := selectFolderEx("", "Select a folder containing " arrayToComma(selectedInputExt) " type files.", winExist(mainAppName)) ;fileSelectFolder, inputFolder, % "::{20d04fe0-3aea-1069-a2d8-08002b30309d}", 3, % "Select a folder containing " extFolder "type files."  
			if ( inputFolder.SelectedDir ) {
				inputFolder := regExReplace(inputFolder.SelectedDir, "\\$")
				for idx, thisExt in selectedInputExt
					loop, Files, % inputFolder "\*." thisExt, FR 
						newFiles.push(a_LoopFileLongPath)
			}
	}
	
	if ( newFiles.length() ) {
		gui 1: listView, listViewInputFiles
		for idx, thisFile in newFiles {
			if ( !inArray(thisFile, scannedFiles[jobCmd]) ) {
				numAdded++
				log("Adding '" thisFile "'")
				SB_SetText("Adding '" thisFile "'", 1, 1)
				LV_add("", thisFile)
				scannedFiles[jobCmd].push(thisFile)
			} else  {
				log("Skip adding '" thisFile "' - already in list.")
				SB_SetText("Skip adding '" thisFile "' - already in list.", 1)
			}
		}
		log("Added " numAdded " files")
	}

	if ( scannedFiles[jobCmd].length() ) {
		;guiToggle("disable", ["buttonInputExtSelect", "buttonOutputExtSelect"])
		log(scannedFiles[jobCmd].length() " jobs in the " stringUpper(jobCmd) " queue. Ready to start!")
	}
	refreshGUI()
}


; Listview containting input files was clicked
; --------------------------------------------
listViewInputFiles()
{
	global
	local suffx, idx, val
	
	if ( !a_eventInfo )
		return

	if ( LV_GetCount("S") > 0 )  {
		suffx := (LV_GetCount("S") > 1) ? "s" : ""
		guiCtrl({buttonRemoveInputFiles:"Remove file" suffx, buttonclearInputFiles:"Clear selection" suffx})
		for idx, val in [4,6]																					; Apply buttons 'Remove selections' and 'Clear selections' skin after changing text
			imageButton.create(GUIbutton%val%, GUI.buttons.default.normal, GUI.buttons.default.hover, GUI.buttons.default.clicked, GUI.buttons.default.disabled) ; Default button colors
		guiToggle("enable", ["buttonRemoveInputFiles", "buttonclearInputFiles"])
	}
	else
		guiToggle("disable", ["buttonRemoveInputFiles", "buttonclearInputFiles"])

}


; Select from input listview files
; --------------------------------
selectInputFiles()
{
	global
	local row := 0, removeThese := [], removeThisFile
	
	gui 1:submit, nohide
	gui 1:+OwnDialogs
	gui 1: listView, listViewInputFiles 					; Select the listview for manipulation
	
	if ( inStr(a_GuiControl, "Select") )
		LV_Modify(0, "Select")
	else if ( inStr(a_GuiControl, "Clear") )
		LV_Modify(0, "-Select")
	else if ( inStr(a_GuiControl, "Remove") ) {
		guiToggle("disable", ["buttonclearInputFiles", "buttonSelectAllInputFiles", "buttonRemoveInputFiles", "buttonAddFolder", "buttonAddFiles"])
		loop {
			row := LV_GetNext(row)										; Get selected download from list and move to next
			if ( !row ) 												; Break if no more selected
				break
			removeThese.push(row)
		}
		while ( removeThese.length() ) {
			row := removeThese.pop()
			LV_GetText(removeThisFile, row , 1)
			LV_Delete(row)
			removeFromArray(removeThisFile, scannedFiles[jobCmd])
			log("Removed '" removeThisFile "' from queue")
			SB_SetText("Removed '" removeThisFile "' from queue" , 1)
		}
	}
	log( scannedFiles[jobCmd].length()? scannedFiles[jobCmd].length() " jobs in the " stringUpper(jobCmd) " queue. Ready to start!" : "No jobs in the " stringUpper(jobCmd) " file queue" )
	controlFocus, SysListView321, %mainAppName%
}


; Select output folder
; --------------------
editOutputFolder()
{
	global
	local newFolder, badChar, folderChk
	
	gui 1:submit, nohide
	gui 1:+OwnDialogs
	
	newFolder := editOutputFolder
	if ( a_guiControl == "buttonBrowseOutput" ) {
		newFolder := selectFolderEx(outputFolder, "Select a folder to save converted files to", mainAppHWND)
		newFolder := newFolder.selectedDir
	}
	if ( !newFolder || newFolder == outputFolder ) 
		return
	for idx, val in ["*", "?", "<", ">", "/", "|", """"]
		badChar := inStr(newfolder, val) ? true : false
	folderChk := splitPath(newFolder)
	if ( !folderChk.drv || !folderChk.dir || badChar ) 			; Make sure newFolder is a valid directory string
		msgBox % "Invalid output folder"
	else {
		outputFolder := regExReplace(newFolder, "\\$")
		log("'" outputFolder "' selected as output folder")
		SB_SetText("'" outputFolder "' selected as new output folder" , 1)
		ini("write", "outputFolder")
		refreshGUI()
		controlFocus,, %mainAppName%
	}
	guiCtrl({editOutputFolder:outputFolder})	; Replace edit field with new outputfolderor reverts back to old value if new value invalid
}



; An chdmanOpt checkbox was clicked
; ---------------------------------
checkboxOption()
{
	global
	local opt
	
	gui 1:submit, nohide
	opt := strReplace(a_guicontrol, "_checkbox")
	guiToggle((%a_guiControl%? "enable":"disable"), [opt "_dropdown", opt "_edit"])
	if ( chdmanOpt[opt].hasKey("masterOf") ) {	; Disable 'slave' checkbox if set in options
		guiToggle((%a_guiControl%? "enable":"disable"), [chdmanOpt[opt].masterOf "_checkbox", chdmanOpt[opt].masterOf "_dropdown", chdmanOpt[opt].masterOf "_edit"])
		guiCtrl({(chdmanOpt[opt].masterOf "_checkbox"):0})
	}
}



; Convert job button - Start the jobs!
; ------------------------------------
buttonStartJobs()
{
	global
	local fnMsg, runCmd, thisJob, gPos, y, dropdownCHDInfoList:="", file, filefull, dir, x1, x2, y
	static CHDInfoFileNum
	gui 1:submit, nohide
	gui 1:+OwnDialogs
	
	switch jobCmd { 
		case "createcd", "createhd", "createraw", "createld", "extractcd", "extracthd", "extractraw", "extractld":
			SB_SetText("Creating work Queue" , 1)
			log("Creating work queue")
			workQueue := createWorkQueue(jobCmd, jobOptions, selectedOutputExt, scannedFiles[jobCmd], regExReplace(outputFolder, "\\$"))	; Create a queue (object) of files to process
		
		case "verify":
			workQueue := createWorkQueue("verify", "", "", scannedFiles["verify"])
		
		case "info":
			CHDInfoFileNum := 1, x1 := 20, x2:= 150, y := 20
			gui 3: destroy
			gui 3: margin, 20, 20
			gui 3: font, s12 Q5 w700 c000000
			gui 3: add, text, x20 y%y% w500 vtextCHDInfoTitle, % ""
			y += 40
			loop 12 {
				gui 3: font, s9 Q5 w700 c000000
				gui 3: add, text, x%x1% y%y% w100 vtextCHDInfoTextInfoTitle_%a_index%, % ""
				gui 3: font, s9 Q5 w400 c000000
				gui 3: add, edit, x%x2% y+-13 w615 veditCHDInfoTextInfo_%a_index% readonly, % ""
				y += 23
			}
			y += 30
			gui 3: font, s9 Q5 w700 c000000
			gui 3: add, text, x%x1% y%y%, % "Hunks"
			gui 3: add, text, x130 y%y%, % "Type"
			gui 3: add, text, x260  y%y%, % "Percent"
			gui 3: font, s9 Q5 w400 c000000
			loop 4 {
				y += 20
				gui 3: add, text, x%x1% y%y% w150 vtextCHDInfoHunks_%a_index%, % ""
				gui 3: add, text, x130 y%y% w150 vtextCHDInfoType_%a_index%, % ""	; recombine words that were seperated with strSplit  
				gui 3: add, text, x260 y%y% w150 vtextCHDInfoPercent_%a_index%, % ""
			}
			y += 40
			gui 3: font, s9 Q5 w700 c000000
			gui 3: add, text, x%x1% y%y% w150, % "Metadata"
			gui 3: font, s8 Q5 w400 c000000, % "Consolas"
			y += 20
			gui 3: add, edit, x%x1% y%y% w750 h250 veditMetadata readonly, % ""
			y+= 270
			gui 3: font, s9 Q5 w400 c000000
			gui 3: add, button, x20 w120 y%y% h30 gselectCHDInfo vbuttonCHDInfoLeft hwndbuttonCHDInfoPrevHWND disabled, % "< Prev file"
			imageButton.create(buttonCHDInfoPrevHWND, GUI.buttons.default.normal, GUI.buttons.default.hover, GUI.buttons.default.clicked, GUI.buttons.default.disabled) ; Default button colors
			
			for idx, filefull in scannedFiles["info"]
				dropdownCHDInfoList .= splitPath(filefull).file "|" ; Create CHD info dropdown list
			y+=5
			gui 3: add, dropdownlist, x160 y%y% w475 vdropdownCHDInfo gselectCHDInfo altsubmit, % regExReplace(dropdownCHDInfoList, "\|$")
			y-=5
			gui 3: font, s9 Q5 w400 c000000
			gui 3: add, button, x+15 y%y% w120 h30 gselectCHDInfo vbuttonCHDInfoRight hwndbuttonCHDInfoNextHWND,  % " Next file >"
			imageButton.create(buttonCHDInfoNextHWND, GUI.buttons.default.normal, GUI.buttons.default.hover, GUI.buttons.default.clicked, GUI.buttons.default.disabled) ; Default button colors
			gui 3: +toolWindow
			gui 3:show, autosize, % "CHD Info"

			selectCHDInfo:
				gui 3:submit, nohide
				switch a_guiControl {
					case "buttonCHDInfoLeft":
						CHDInfoFileNum--
					case "buttonCHDInfoRight":
						CHDInfoFileNum++
					case "dropdownCHDInfo":
						CHDInfoFileNum := dropdownCHDInfo
				}
				if ( showCHDInfo(scannedFiles["info"][CHDInfoFileNum]) == false )
					return

				guiCtrl("choose", {dropdownCHDInfo:CHDInfoFileNum}, 3)					; Choose dropdown to match newly selected item in CHD info
				controlFocus, ComboBox1, % "CHD Info"									; Keep focus on dropdown to allow arrow keys to also select next and previous files
				
				guiToggle("enable", ["buttonCHDInfoLeft","buttonCHDInfoRight"], 3)		; Enable both selection buttons by default
				if ( CHDInfoFileNum == 1 )												; Then disable the appropriate button according to selection number (first or last in list)
					guiToggle("disable", "buttonCHDInfoLeft", 3)						
				else if ( CHDInfoFileNum == scannedFiles["info"].length() )
					guiToggle("disable", "buttonCHDInfoRight", 3)
				
			return

		case "addMeta":
		case "delMeta":
		case default:
	}
	if ( workQueue.length() == 0 )  															; Nothing to do!
		return
	
	msgData := [], thisJob := {}, availPSlots := []
	jobWorkTally := {started:false, total:workQueue.length(), success:0, cancelled:0, skipped:0, withError:0, finished:0, haltedMsg:"", report:""}		; Set variables
	workQueueSize := (jobWorkTally.total < jobQueueSize)? jobWorkTally.total : jobQueueSize	 								; If number of jobs is less then queue count, only display those progress bars

	dllCall("SetMenu", "uint", mainAppHWND, "uint", 0), mainMenuVisible := false											; Hide main menu bar (selecting menu when running jobs stops messages from being receieved from threads)								
	guiToggle("disable", "all")																								; Disable all controls while job is in progress
	guiCtrl({buttonStartJobs:"CANCEL ALL JOBS"})																			; Rename Start button to cancel jobs
	guiToggle("enable", "buttonStartJobs")
	imageButton.create(startButtonHWND, GUI.buttons.cancel.normal, GUI.buttons.cancel.hover, GUI.buttons.cancel.clicked)	; Change color of start job button to red
	gPos := (jobCmd == "verify" || jobCmd == "info")? guiGetCtrlPos("groupboxJob") : guiGetCtrlPos("groupboxOptions")		; Move and show progress bars
	y := gPos.y + gPos.h + 25																								; Assign y (x are from 'gPos') values to groupbox x & y positions
	guiCtrl("moveDraw", {progressAll:"y" y, progressTextAll: "y" y+4})														; Set All Progress bar and it's text Y position
	y += 35
	loop % workQueueSize {																		
		guiCtrl("moveDraw", {("progress" a_index):"y" y, ("progressText" a_index):"y" y+4, ("progressCancelButton" a_index):"y" y})	; Move the progress bars into place
		guiCtrl({("progress" a_index):0, ("progressText" a_index): ""})														; Clear the bars text and zero out percentage
		y += 25
	}
	guiCtrl( {progressAll:0, progressTextAll:"0 jobs of " jobWorkTally.total " completed - 0%"})
	guiCtrl("moveDraw", {groupBoxProgress:"x5 y" (gPos.y + gPos.h) + 5 " h" workQueueSize*25 + 60})							; Move and resize progress groupbox
	guiToggle("show", ["groupBoxProgress", "progressAll", "progressTextAll"])												; Show total progress bars					
	loop % workQueueSize {																	
		guiToggle("show", ["progress" a_index, "progressText" a_index, "progressCancelButton" a_index])						; Show job progress bars
		availPSlots.push(a_index)																							; Add available progress slots to queue
	}
	onMessage(0x0201, "clickedGUI")																							; Mouse button was clicked in GUI window
	gui 1:show, autosize																									; Resize main window to fit progress bars
	log(jobWorkTally.total " " stringUpper(jobCmd) " jobs starting ...")
	SB_SetText(jobWorkTally.total " " stringUpper(jobCmd) " jobs started" , 1)
	setTimer, jobTimeoutTimer, 1000																							; Check for timeout of chdman or thread
	jobWorkTally.started := true
	
	loop {
		if ( availPSlots.length() > 0 && workQueue.length() > 0 ) {									; Wait for an available slot in the queue to be added
			thisJob := workQueue.removeAt(1)														; Grab the first job from the work queue and assign parameters to variable
			thisJob.pSlot := availPSlots.removeAt(1)												; Assign the progress bar a y position from available queue
			msgData[thisJob.pSlot] := {}
			msgData[a_index].timeout := 0
			
			runCmd := a_ScriptName " threadMode " (showJobConsole == "yes" ? "console" : "")		; "threadmode" flag tells script to run this script as a thread
			run % runCmd ,,, pid																	; Run it
			thisJob.pid := pid
			
			while ( pid <> msgData[thisJob.pSlot].pid ) {											; Wait for confirmation that msg was receieved												
				sendAppMessage(toJSON(thisJob), "ahk_class AutoHotkey ahk_pid " pid)
				sleep 25
			}
		}
		
		if ( jobWorkTally.finished == jobWorkTally.total || jobWorkTally.started == false )
			break																					; Job queue has finished
	
		sleep 250
	}
	
	setTimer, jobTimeoutTimer, off
	jobWorkTally.started := false
	guiToggle("disable", "all")
	onMessage(0x0201, "")																		; Turn off mouse button was clicked in GUI window
	
	if ( jobWorkTally.haltedMsg ) {																	; There was a fatal error that didnt allow any jobs to be attempted
		log("Fatal Error: " jobWorkTally.haltedMsg)
		SB_SetText("Fatal Error: " jobWorkTally.haltedMsg , 1)
		msgBox, 16, % "Fatal Error", jobWorkTally.haltedMsg "`n"
	}
	else {																						; {normal:0, cancelled:0, skipped:0, withError:0, halt:0}
		fnMsg := "Total number of jobs attempted: " jobWorkTally.total "`n"
		fnMsg .= jobWorkTally.success ? jobFinPreTxt " sucessfully: " jobWorkTally.success "`n" : ""
		fnMsg .= jobWorkTally.cancelled ? "Jobs cancelled by the user: " jobWorkTally.cancelled "`n" : ""
		fnMsg .= jobWorkTally.skipped ? "Jobs skipped because the output file already exists: " jobWorkTally.skipped "`n" : ""
		fnMsg .= jobWorkTally.withError ? "Jobs that finished with errors: " jobWorkTally.withError : ""
	
		SB_SetText("Jobs finished" (jobWorkTally.withError? " with some errors":"!"), 1)
		log( regExReplace(strReplace(fnMsg, "`n", ", "), ", $", "") )
		
		if ( playFinishedSong == "yes" && jobWorkTally.success )												; Play sounds to indicate we are done (only if at least one successful job)
			playSound()		
	
		msgBox, 4, % mainAppName, % "Finished jobs!`nWould you like to see a report?"
		ifMsgBox Yes
		{
			gui 3: destroy
			gui 3: margin, 10, 20
			gui 3: -sysmenu
			gui 3: font, s11 Q5 w700 c000000
			gui 3: add, text,, % jobDesc " report"
			gui 3: font, s9 Q5 w400 c000000
			gui 3: add, edit, y+15 w800 h500, % fnMsg "`n`n" jobWorkTally.report
			gui 3: add, button, x350 y+15 w100 h24 gfinishJob, OK								; go to finishJob here
			gui 3: show, autosize center, REPORT
			controlFocus,, REPORT
			return
		}
		else
			finishJob()
	}
}



; All jobs have finished or user pressed okay after report
; --------------------------------------------------------
finishJob()
{		
	global workQueueSize	
	
	guiToggle("hide", ["groupBoxProgress", "progressAll", "progressTextAll"])	; Show total progress bars					
	loop % workQueueSize																	
		guiToggle("hide", ["progress" a_index, "progressText" a_index, "progressCancelButton" a_index])	; Show job progress bars
	gui 3: destroy
	refreshGUI()
}


; Cancel a single job in progress
; -------------------------------
progressCancelButton()
{
	global msgData
	
	if ( !a_guiControl )
		return
	pSlot := strReplace(a_guiControl, "progressCancelButton", "")
	msgBox, 4,, % "Cancel job " msgData[pSlot].idx " - " stringUpper(msgData[pSlot].cmd) ": " msgData[pSlot].workingTitle "?", 15
	ifMsgBox Yes
		cancelJob(pSlot)
}




; Create the main GUI
; -------------------
createMainGUI() 
{
	global
	local thisOptName, key, thisOpt, thisBtn
	
	gui 1:add, button, 		hidden default h0 w0 y0 y0 geditOutputFolder,		; For output edit field (default button
	
	gui 1:add, statusBar
	SB_SetParts(640, 175)
	SB_SetText("  namDHC v" mainAppVersion " for CHDMAN", 2)

	gui 1:add, groupBox, 	x5 w800 h425 vgroupboxJob, Job

	gui 1:add, text, 		x15 y30, % "Job type:"
	gui 1:add, dropDownList,x+5 y28 w200 vdropdownJob gselectJob, % GUI.dropdowns.job.create "||" GUI.dropdowns.job.extract "|" GUI.dropdowns.job.info "|" GUI.dropdowns.job.verify 	;"|" GUI.dropdowns.job.addMeta "|" GUI.dropdowns.job.delMeta

	gui 1:add, text, 		x+30 y30, % "Media type:"
	gui 1:add, dropDownList,x+5 y28 w200 vdropdownMedia gselectJob, % GUI.dropdowns.media.cd "||" GUI.dropdowns.media.hd "|" GUI.dropdowns.media.ld "|" GUI.dropdowns.media.raw

	gui 1:add, text, 		x15 y65, % "Input files"
	
	gui 1:add, button, 		x15 y83 w80 h22 vbuttonAddFiles gaddFolderFiles hwndGUIbutton2, % "Add files"
	gui 1:add, button, 		x+5 y83 w90 h22 vbuttonAddFolder gaddFolderFiles hwndGUIbutton3, % "Add a folder"
	
	gui 1:add, text, 		x475 y93, % "Input file types: "
	gui 1: font, Q5 s9 w700 c000000
	gui 1:add, text, 		x+3 y93 w110 vselectedInputExtText, % ""
	gui 1: font, Q5 s9 w400 c000000
	gui 1:add, button,		x663 y83 w130 h22 gbuttonExtSelect vbuttonInputExtSelect hwndGUIbutton1, % "Select input file types"
	
	gui 1:add, listView, 	x15 y110 w778 h153 vlistViewInputFiles glistViewInputFiles altsubmit, % "File"
	
	gui 1:add, button, 		x15 y267 w90 vbuttonSelectAllInputFiles gselectInputFiles hwndGUIbutton5, % "Select all"
	gui 1:add, button, 		x+5 y267 w90 vbuttonClearInputFiles gselectInputFiles hwndGUIbutton6, % "Clear selection"
	gui 1:add, button, 		x+20 y267 w90 vbuttonRemoveInputFiles gselectInputFiles hwndGUIbutton4, % "Remove selection"

	gui 1:add, text, 		x15 y305, % "Output Folder"
	gui 1:add, button, 		x15 y324 w90 vbuttonBrowseOutput geditOutputFolder hwndGUIbutton8, % "Select a folder"
	gui 1:add, text, 		x485 y335,% "Output file type: "
	gui 1: font, Q5 s9 w700 c000000
	gui 1:add, text, 		x+3 y335 w100 vselectedOutputExtText, % ""
	gui 1: font, Q5 s9 w400 c000000
	gui 1:add, button,		x663 y324 w130 h24 vbuttonOutputExtSelect gbuttonExtSelect hwndGUIbutton7, % "Select output file type"
	
	gui 1:add, edit, 		x15 y352 w778 veditOutputFolder +wantReturn, % outputFolder
	
	gui 1:add, button,		x320 y385 w160 h35 vbuttonStartJobs gbuttonStartJobs hwndstartButtonHWND, % "Start all jobs!"

	gui 1:add, groupBox, 	x5 w800 y435 vgroupboxOptions, % "CHDMAN Options"		; Position and height will be set in refreshGUI()

	loop 8 {	; Stylize default buttons
		thisBtn := "GUIbutton" a_index
		imageButton.create(%thisBtn%, GUI.buttons.default.normal, GUI.buttons.default.hover, GUI.buttons.default.clicked, GUI.buttons.default.disabled) ; Default button colors
	}
	
	for key, thisOpt in chdmanOpt
	{
		; Options are moved to their positions when refreshGUI(true) is called
		; -------------------------------------------------------------------
		if ( thisOpt.hidden == true )
			continue
		thisOptName := thisOpt.name
		gui 1:add, checkbox,		hidden w200 gcheckboxOption -wrap v%thisOptName%_checkbox,
		gui 1:add, edit,			hidden w165 v%thisOptName%_edit,
		gui 1:add, dropdownList, 	hidden w165 altsubmit v%thisOptName%_dropdown,		; ... so we can use for dropdown list to place at same location (default is hidden)
	}
}


; Create GUI progress bar section
; -------------------------------
createProgressBars()
{
	global
	local thisBtn
	
	gui 1:add, groupBox, w800 vgroupBoxProgress, Progress

	gui 1: font, Q5 s9 w700 cFFFFFF
	gui 1:add, progress, hidden x20 w770 h22 backgroundAAAAAA vprogressAll cgreen, 0		; Progress bars y values will be determined with refreshGUI()
	gui 1:add, text,	 hidden x30 w750 h22 +backgroundTrans -wrap vprogressTextAll

	loop % jobQueueSizeLimit {																; Draw but hide all progress bars - we will only show what is called for later
		gui 1:add, progress, hidden x20 w740 h22 backgroundAAAAAA vprogress%a_index% c17A2B8, 0				
		gui 1:add, text,	 hidden x30 w720 h22 +backgroundTrans -wrap vprogressText%a_index%
		gui 1:add, button,	 hidden x+15 w25 vprogressCancelButton%a_index% gprogressCancelButton hwndprogCancelbutton%a_index%, % "X"
		thisBtn := "progCancelbutton" a_index
		imageButton.create(%thisBtn%, GUI.buttons.cancel.normal,	GUI.buttons.cancel.hover, GUI.buttons.cancel.clicked)
	}
}


createMenus() 
{
	global GUI, jobQueueSizeLimit, jobQueueSize

	obj := GUI.menu
	
	loop % jobQueueSizeLimit
		menu, SubSettingsConcurrently, Add, %a_index%, % "menuSelected"
	menu, SubSettingsConcurrently, Check, % jobQueueSize						; Select current jobQueue number

	loop % obj.namesOrder.length() {
		menuName := obj.namesOrder[a_index]
		menuArray := obj[menuName]
		
		loop % menuArray.length() {
			menuItem :=  menuArray[a_index]
			menu, % menuName "Menu", Add, % menuItem.name, % menuItem.gotolabel
		
			if ( menuItem.saveVar ) {
				saveVar := menuItem.saveVar
				menu, % menuName "Menu", % (%saveVar% == "yes"? "Check":"UnCheck"), % menuItem.name
			}
		}
		menu, MainMenu, Add, % menuName, % ":" menuName "Menu"
	}
	gui 1:menu, MainMenu
}


; Refreshes GUI to reflect current settings or to default (clear) with true param
; -------------------------------------------------------------------------------
refreshGUI(resetGUI:=false) {
	global
	local opt, key, val, idx, optNum, checkOpt, changeOpt, x, y, yH, gPos
	static selectedJob
	gui 1:submit, nohide

	guiToggle("enable", ["dropdownJob", "dropdownMedia", "buttonExtSelect", "listViewInputFiles", "buttonAddFiles", "buttonAddFolder", "editOutputFolder", "buttonOutputExtSelect", "buttonBrowseOutput", "buttonStartJobs"])																	; Renable all elements
	guiToggle("disable", ["buttonRemoveInputFiles", "buttonClearInputFiles", "buttonSelectAllInputFiles"])
	guiToggle("hide", "groupBoxProgress") 														; By default we hide progress bars and groupbox - startjob label will show them
	guiToggle((scannedFiles[jobCmd].length()? "enable":"disable"), ["buttonStartJobs", "buttonSelectAllInputFiles"]) 
	guiCtrl({buttonStartJobs:"START JOBS"})														; Start button to default label
	imageButton.create(startButtonHWND, GUI.buttons.start.normal, GUI.buttons.start.hover, GUI.buttons.start.clicked, GUI.buttons.start.disabled) ; Default button colors,  must be set after changing button text

	; Refresh GUI to defaults
	if ( resetGUI ) { 																			; Only hide & clear all checkboxes, editboxes and dropdowns if changing jobs or media types
		for key, opt in chdmanOpt {
			guiToggle("hide", [opt.name "_checkbox", opt.name "_edit", opt.name "_dropdown"])	; Hide all checkboxes, editfields and dropdowns	
			guiCtrl({(opt.name "_checkbox"):0})													; Uncheck all checkboxs
		}
		
		if ( jobOptions.length() < 1 )	{															
			guiToggle("hide", "groupboxOptions")												; No options to show for this job, so hide the group
			guiCtrl("moveDraw", {groupboxOptions:"h0"})											; Resize options groupbox height for no options
		}
		else {																					; Show or hide checkbox options
			guiToggle("show", "groupboxOptions")
			
			gPos := guiGetCtrlPos("groupboxOptions")											; Assign x & y values to groupbox x & y positions
			idx := 0, yH := 0, x := gPos.x+10, y := gPos.y										
			
			for optNum, opt in jobOptions { 													; Show appropriate options according to selected type of job and media type
				if ( !opt || opt.hidden )
					continue
				
				if ( opt.hasKey("editField") )													; The option can ONLY have either a dropdown or editfield
					changeOpt := opt.name "_edit"
				else if ( opt.hasKey("dropdownOptions") )
					changeOpt := opt.name "_dropdown"
				checkOpt := opt.name "_checkbox"

				if ( idx == chdmanOptionMaxPerSide )											; We've run out of vertical room to show more chdman options so new options go into next column
					x += 400, y := gPos.y, yH := 1, idx := 0									
				else 
					yH++, idx++
				
				guiCtrl({(changeOpt):inStr(changeOpt,"dropdown") ? opt.dropdownOptions : opt.editField})	; Populate option dropdown list from chdmanOpt array	
				guiCtrl({(checkOpt):" " (opt.description? opt.description : opt.name)}) 					; Label the checkbox -- Default to using the chdman opt.name parameter if no description
				guiCtrl("moveDraw", {(checkOpt):"x" x + (opt.xInset? opt.xInset:0) " y" y+(yH*25), (changeOpt):"x" x+210 " y" y +(yH*25)-3})	; Move the chdman option editbox or dropdown control into place and Move the checkbox into place
				guiToggle("show", [checkOpt, changeOpt])													; Show the option and its coresponding editfield or dropdownlist
			}
			guiCtrl("moveDraw", {groupboxOptions:"h" ceil((optNum >9 ? 9 : optNum)*25)+30})					; Resize options groupbox height to fit all  options
		}
		
		; Populate extension menu lists
		for idx2, type in ["Input", "Output"] {													; Populate extension dropdown lists	
			drpList := "job" type "Exts"
			buildList := "selected" type "Ext"  												; 'selectedInputExt' and 'selectedOutputExt'
			if ( %buildList% ) 
				menu, % type "ExtTypes", deleteAll		 										; Delete previous menu if it exists
			%buildList% := []
			for idx, ext in %drpList% {
				if ( ext ) {
					menu, % type "ExtTypes", add, % ext, % "menuExtHandler"						; Add extension to the menu
					if ( dropdownJob == GUI.dropdowns.job.extract && idx > 1 )					; By default, only check one (1st) extension if we are extracting an image
						continue
					else {
						menu, % type "ExtTypes", Check, % ext									; Check all extension menu items unless we are extracting
						%buildList%.push(ext)													; Add it to input/output ext menu
					}
				}
			}
			guiCtrl({(buildList) "Text": arrayToComma(%buildList%)})
		}

		; Populate listview
		gui 1: listView, listViewInputFiles														; Select the main file listview										
		LV_delete()																				; Delete all entries
		for idx, thisFile in scannedFiles[jobCmd]
			LV_add("", thisFile)																; Re-populate listview with scanned files
		controlFocus, SysListView321, %mainAppName%												; Focus on listview  to stop one item being selected
	} ; End resetGUI
	
	; Enable or disable editfields or dropdown boxes depening  on checked status
	for optNum, opt in jobOptions {																
		checkOpt := opt.name "_checkbox"
		guiToggle((%checkOpt%? "enable":"disable"), [opt.name "_dropdown", opt.name "_edit", opt.masterOf "_checkbox", opt.masterOf "_dropdown", opt.masterOf "_edit"]) ; If checked, enable or disable dropdown or editfields
	}
	
	; Other changes depending on job selected
	switch dropdownJob {																						; variable from dropdown gui
		case GUI.dropdowns.job.create, GUI.dropdowns.job.extract:
			guiToggle("enable", ["listViewInputFiles", "dropdownMedia", "editOutputFolder", "buttonBrowseOutput"])
			if ( selectedJob <> dropdownJob ) {																	; Stop recursive selection												
				guiCtrl({dropdownMedia:"|" GUI.dropdowns.media.cd "|" GUI.dropdowns.media.hd "|" GUI.dropdowns.media.ld "|" GUI.dropdowns.media.raw})
				guiCtrl("choose", {dropdownMedia:"|1"}) 														; Choose first item in media dropdown and fire the selection
			}
		case GUI.dropdowns.job.info, GUI.dropdowns.job.verify:
			guiToggle("disable", ["dropdownMedia", "buttonOutputExtSelect", "buttonInputExtSelect", "editOutputFolder", "buttonBrowseOutput"])
			guiCtrl({dropdownMedia:"|CHD Files"})
			guiCtrl("choose", {dropdownMedia:1})																; Choose first item in media dropdown but since its disabled, dont fire the selection

		case GUI.dropdowns.job.addMeta, GUI.dropdowns.job.delMeta:
			guiToggle("disable", "all")
			guiToggle("enable", "dropdownJob")
			msgbox % "Not implemented yet"
	}
	selectedJob := dropdownJob
	
	; Show main menu
	if ( !mainMenuVisible && mainAppMenuGet ) {
		dllCall("SetMenu", "uint", mainAppHWND, "uint", mainAppMenuGet) 					
		mainAppMenuGet := DllCall("GetMenu", "uint", mainAppHWND)
		mainMenuVisible := true
	}
	
	; Set status text
	SB_SetText(scannedFiles[jobCmd].length()? scannedFiles[jobCmd].length() " jobs in the " stringUpper(jobCmd) " queue. Ready to start!" : "Add files to the job queue to start", 1)
	
	; Show and resize main GUI
	gui 1:show, autosize x%mainWinPosX% y%mainWinPosY%, % mainAppName
}




; Show or hide the verbose window
; -------------------------------
showVerboseWindow(show:="yes")
{
	global
	static created
	
	if ( !created ) {
		created := true
		gui 2:-sysmenu +resize
		gui 2:margin, 5, 10
		gui 2:add, edit, % "w" verboseWinPosW-10 " h" verboseWinPosH-20 " readonly veditVerbose",
	}
	if ( show == "yes" ) {
		gui 2:show, % "w" verboseWinPosW " h" verboseWinPosH " x" verboseWinPosX " y" verboseWinPosY, % mainAppNameVerbose
		sendMessage 0x115, 7, 0, Edit1, % mainAppNameVerbose		; Scroll to bottom of log
		controlFocus,, % mainAppNameVerbose
	}
	else if ( show == "no" ) {
		gui 2:hide
	}
}


; Show CHD info info seperate window
; -- grab new data 'JIT'
; ----------------------------------
showCHDInfo(fullFileName)
{
	global
	local a, line, file, infoLineNum := 0, compressLineNum := 0, metadataTxt := ""
	
	if ( !fileExist(fullFileName) )
		return false
		
	file := splitPath(fullFilename)
	guiCtrl({"textCHDInfoTitle":file.file}, 3) 																	; Change Title to filename
	loop, parse, % runCMD(chdmanLocation " info -v -i """ fullFilename """", file.dir).msg, % "`n"				; Loop through chdman 'info' stdOut
	{
		if ( a_index == 1 )																						; Skip first line of output
			continue
		line := regExReplace(a_loopField, "`n|`r")																; Remove all CR/LF
		if ( inStr(line, "Metadata") ) {																		; If we find string 'Metadata' in line, we know to add text to metadata string
			line := strReplace(line, "Metadata:")																; Remove 'Metadata:' as its redundant
			metadataTxt .= trim(line, " ") "`n"
		}
		if ( inStr(line, "TRACK:") ) {																			; Finding 'TRACK:' in text informs us we are in metadata section of output
			line := trim(line, " "), line := strReplace(line, " ", " | "), line := strReplace(line, ":", ": ")  ; Fix formatting ...
			metadataTxt .= strReplace(line, ".") "`n`n"															; ... and add it to the metadata string
		}
		else if ( inStr(line, ": ") ) {																			; Otherwise all data is part of file information
			infoLineNum++																						; Increase line number counter
			a := strSplit(line, ": ")																			; Split text into parts
			guiCtrl({("textCHDInfoTextInfoTitle_" infoLineNum):trim(a[1], " ") ": "}, 3)						; Add part 1 as subtitle (ie - "File name", "Size", "SHA1", etc)
			guiCtrl({("editCHDInfoTextInfo_" infoLineNum):trim(a[2], " ")}, 3)									; Part 2 is the information itself 
		}
		else if ( line == "----------  -------  ------------------------------------" )								; When we find this string, we know we are in the overview Compression section
			compressLineNum := 1																					; ... So flag it, use flag as the line counter and move to next loop
		else if ( compressLineNum ) {																					
			line := trim(line, a_space)
			line := regExReplace(line, "      |     |    |   |  ", ";")												; Change all "|" into ";" in line and remove redundant space
			a := strSplit(line, ";")																				; Then split it into part
			if ( a[1] ) {
				guiCtrl({("textCHDInfoHunks_" compressLineNum):trim(a[1], a_space)}, 3)								; Part 1 is Hunks
				guiCtrl({("textCHDInfoType_" compressLineNum):trim(a[3] " " a[4] " " a[5] " " a[6], a_space)}, 3)	; Part 2 is Compression Type
				guiCtrl({("textCHDInfoPercent_" compressLineNum):trim(a[2], a_space)}, 3)							; Part 3 is percentage of compression
			}
			compressLineNum++																						; Add to meta line number
		}
	}
	guiCtrl({"editMetadata": metadataTxt}, 3)
	controlFocus, , % "CHD Info"
	return true
}



; guiToggle GUI controls
; -------------------------------------------------------
guiToggle(doWhat, whichControls, guiNum=1) 
{
	global mainAppName
	ctlArray := [], doWhatArray := []
	
	if ( isObject(doWhat) )
		doWhatArray := doWhat
	else if ( doWhat <> "" )
		doWhatArray[1] := doWhat
	else return false	
	
	if ( isObject(whichControls) )
		ctlArray := whichControls
	else if ( whichControls <> "" )
		ctlArray[1] := whichControls
	else return false
	
	if ( ctlArray[1] == "all" ) {
		for idx, dw in doWhatArray {
			;winGet, ctrList, ControlList, % mainAppName
			;ctlArray := 
			for idx, ctl in ["dropdownJob", "dropdownMedia", "buttonExtSelect", "listViewInputFiles", "buttonAddFiles", "buttonAddFolder", "buttonRemoveInputFiles", "buttonselectAllInputFiles", "buttonclearInputFiles", "editOutputFolder", "buttonOutputExtSelect", "buttonBrowseOutput", "buttonStartJobs"]
			{	
				if ( dw == "toggle" ) {
					guiControlGet, vis, %guiNum%:Visible, % ctl
					dw := vis ? "hide" : "show"
				}
				guiControl %guiNum%:%dw%, % ctl
			}
		}
	}
	else {
		for idx, dw in doWhatArray {
			for idx2, ctl in ctlArray  {
				if ( dw == "toggle" ) {
					guiControlGet, vis, %guiNum%:Visible, % ctl
					dw := vis ? "hide" : "show"
				}
				guiControl %guiNum%:%dw%, % ctl
			}
		}
	}
}

; Receieve message data from thread script
; ----------------------------------------
receiveData(wParam, ByRef lParam) 
{
	global msgData, queuedData
	
	data := fromJSON(strGet( numGet(lParam + 2*A_PtrSize) ,, "utf-8"))
	msgData[data.pSlot] := data			; Assign globally so we can use anywhere in script - mainly to kill job if user selects	
	queuedData.push(data)
	
	parseData()
}
	
; Parse data receieved from thread script 
; --------------------------------------
parseData()  					; This is split from receieveData so parsing can be called in other parts of the script without having to receive data from thread
{
	global 
	local recvData, percentAll
	static report := []
	
	while ( queuedData.length() > 0 ) {
		recvData := queuedData.removeAt(1)
		
		if ( recvData.log )
			log("Job " recvData.idx " - " recvData.log)
		
		if ( recvData.report )				
			report[recvData.idx] .= recvData.report		; Static variable adds to end report data

		switch recvData.status {
			case "starting":
				jobWorkTally.started++
			
			case "success":
				jobWorkTally.success++
				SB_SetText("Job " recvData.idx " finished successfully!", 1)

			case "fileExists":
				jobWorkTally.skipped++
				SB_SetText("Job " recvData.idx " skipped", 1)
			
			case "error":
				jobWorkTally.withError++
				SB_SetText("Job " recvData.idx " failed", 1)
				
			case "killed":
				jobWorkTally.cancelled++
				SB_SetText("Job " recvData.idx " cancelled", 1)
			
			case "halted":
				jobWorkTally.cancelled += workQueue.length() + 1				; Tally up totals
				workQueue := []											; Empty the work queue
				jobWorkTally.haltedMsg := recvData.log						; Set flag and error log
				log("Fatal Error. Halted all jobs")
				
			case "finished":
				jobWorkTally.finished++
				jobWorkTally.report .= report[recvData.idx]
				msgData[recvData.pSlot] := ""
				report[recvData.idx] := ""

				percentAll := ceil((jobWorkTally.finished/jobWorkTally.total)*100)
				guiCtrl({progressAll:percentAll, progressTextAll:jobWorkTally.finished " jobs of " jobWorkTally.total " completed " (jobWorkTally.withError ? "(" jobWorkTally.withError " error" (jobWorkTally.withError>1? "s)":")") : "")" - " percentAll "%" })
				if ( removeFileEntryAfterFinish == "yes" ) {
					removeFromArray(recvData.fromFileFull, scannedFiles[recvData.cmd])
					loop % LV_GetCount()												; Clear finished files from scanned files
						if ( LV_GetText2(a_index) == recvData.fromFileFull )
							LV_Delete(a_index)
				}
				availPSlots.push(recvData.pSlot)										; Add an available slot to array
		}
		
		if ( recvData.progress )
			guiControl,1:, % "progress" recvData.pSlot, % recvData.progress
		if ( recvData.progressText )	
			guiControl,1:, % "progressText" recvData.pSlot, % recvData.progressText
	}
}


; Job timeout timer
; Timer is set to call this function every 1000 ms
; ----------------------
jobTimeoutTimer() 
{
	global msgData, jobQueueSize, jobTimeoutSec, queuedData
	
	loop % jobQueueSize {			 	; Loop though jobs to check to see whos msgData is empty - hense to response from thread script
		if ( !msgData[a_index] )		; If data exists return 
			continue
		else
			msgData[a_index].timeout++	; Otherwise add 1 to timer counter
		
		if ( msgData[a_index].timeout >= jobTimeoutSec ) { 			; If timer counter exceeds threashold, we will assume thread is locked up or has errored out 
			processPIDClose(msgData[a_index].chdmanPID, 5, 150)		; So attempt to close the process associated with it
			
			msgData[a_index].status := "error"						; Update msgData[] with messages and send "error" flag for that job, then parse the data
			msgData[a_index].log := "Error: Job timed out"
			msgData[a_index].report := "`nError: Job timed out`n`n`n"
			msgData[a_index].progress := 100
			msgData[a_index].progressText := "Timed out -  " msgData[a_index].workingTitle
			queuedData.push(msgData[a_index])
			parseData()
			
			msgData[a_index].log := ""								; Update msgData[] again, but now send "finished" flag and parse again 
			msgData[a_index].report := ""
			msgData[a_index].status := "finished"
			queuedData.push(msgData[a_index])
			parseData()
			return true
		}
	}
	return false
}




; Create  or add to the input files queue (work queue)
; -------------------------------------------------------
createWorkQueue(command, theseJobOpts, outputExts="", inputFiles="", outputFolder="") 
{
	global
	local idx, thisOpt, optVal, cmdOpts := "", fromFileFull, fileFull, toExt
	local wQueue := []
	
	gui 1:submit, nohide
	
	for idx, thisOpt in (isObject(theseJobOpts) ? theseJobOpts : []) 								; Parse through supplied Options associated with job
	{
		if ( guiCtrlGet(thisOpt.name "_checkbox", 1) == 0 )   											; Skip if the checkbox is not checked
			continue
		if ( thisOpt.editField )
			optVal := guiCtrlGet(thisOpt.name "_edit")
		else if ( thisOpt.dropdownOptions ) {
			optVal := guiCtrlGet(thisOpt.name "_dropdown")											; Get the dropdown value for the current chdmanOpt
			optVal := isObject(thisOpt.dropdownValues) ? thisOpt.dropdownValues[optVal] : optVal	; If this dropdown option contains a dropdownValues array, optVal becomes the index for that array
		}
		if ( thisOpt.paramString ) {
			optVal := optVal ? (thisOpt.useQuotes ? " """ optVal """" : " " optVal) : ""
			cmdOpts .= " -" thisOpt.paramString . optVal 											; Create the chdman options string
		}
	}
	
	for idx1, fromFileFull in (isObject(inputFiles) ?  inputFiles : [])
	{
		fileFull := splitPath(fromFileFull)
		outputExts := isObject(outputExts) ? outputExts : ["dummy"]
		for idx, toExt in outputExts {
			q := {}
			q.idx				:= wQueue.length() + 1
			q.id 				:= command q.idx
			q.hostPID			:= dllCall("GetCurrentProcessId")
			q.cmd 				:= command
			q.cmdOpts			:= cmdOpts
			q.workingDir 		:= fileFull.dir
			q.outputFolder 		:= outputFolder ? outputFolder : ""
			q.fromFile 			:= fileFull.file
			q.fromFileExt		:= fileFull.ext
			q.fromFileNoExt 	:= fileFull.noExt
			q.fromFileFull		:= fromFileFull
			if ( command <> "verify" && command <> "info" ) {
				q.toFile		:= fileFull.noExt "." toExt
				q.toFileExt 	:= toExt
				q.toFileNoExt	:= fileFull.noExt													; For the target file, we use the same base filename as the source
				q.toFileFull	:= outputFolder "\" fileFull.noExt "." toExt
			}
			q.createSubDir		:= createSubDir_checkbox
			q.deleteInputDir	:= deleteInputDir_checkbox
			q.deleteInputFiles 	:= deleteInputFiles_checkbox
			q.keepIncomplete 	:= keepIncomplete_checkbox
			q.workingTitle 		:= (q.toFile ? q.toFile : q.fromFile)
			
			wQueue.push(q)																		; Push data to array
		}
	}
	return wQueue
}


; List filenames from CUE, GDI and TOC files 
; ------------------------------------------
getFilesFromCUEGDITOC(inputFiles) 
{
	fileList := []
	if ( !isObject(inputFiles) )
		inputFiles := Array(inputFiles)
		
	for idx, thisFile in inputFiles {
		if ( !fileExist(thisFile) )
			continue
		f := splitPath(thisFile)
		
		switch f.ext {
		case "cue", "toc":
			loop, Read, % thisFile 
			{
				if ( stPos := inStr(a_loopReadLine, "FILE """, true) ) {
					stPos += 6
					endPos := inStr(a_loopReadLine, """", true, -1)
					file := subStr(a_loopReadLine, stPos, (endPos-stPos))
					fileList.push(f.dir "\" file)
				}
			}
			
		case "gdi":
			loop, Read, % thisFile 
			{
				if ( a_loopReadLine is digit && a_index > 1 ) {
					loop parse, a_loopReadLine, %a_space%
						if ( inStr(a_LoopField, ".") ) 
							fileList.push(f.dir "\" a_loopField)
				}
			}
		}
		fileList.push(thisFile)
	}
	return fileList
}



; Log messages and send to verbose window
; --------------------------------------
log(newMsg:="", newline:=true, clear:=false) 
{
	global mainAppNameVerbose, editVerbose
	
	if ( !newMsg ) 
		return false
	
	newMsg := "[" a_Hour ":" a_Min ":" a_Sec "]  " newMsg
	msg := clear? newMsg : guiCtrlGet("editVerbose", 2) . newMsg

	guiCtrl({editVerbose:msg (newline? "`n" : "")}, 2)
	sendMessage 0x115, 7, 0, Edit1, % mainAppNameVerbose	; Scroll to bottom of log
}


; Mouse button was clicked within gui window
; -------------------------------------------
clickedGUI(wParam, lParam, msg, hwnd)
{
	global jobQueueSize, jobWorkTally, workQueue

	if ( !jobWorkTally.started || a_GuiControl <> "buttonStartJobs" )
		return
		
	if ( jobWorkTally.started == true && a_GuiControl == "buttonStartJobs" ) {
		msgBox, 4,, % "Are you sure you want to cancel all jobs?", 15
		ifMsgBox No 
			return
	
		jobWorkTally.cancelled += workQueue.length()
		jobWorkTally.finished += workQueue.length()
		jobWorkTally.started := false
		workQueue := []													; Clear the work Queue
		loop % jobQueueSize {
			cancelJob(a_index)
		}
	}
}


; User cancels job
; --------------------------------
cancelJob(pSlot)
{
	global msgData
	critical
	
	if ( !pSlot || !msgData[pSlot].pid )
		return

	log("Job " msgData[pSlot].idx " - User requested to cancel...")
	msgData[pSlot].kill := "true"
	return sendAppMessage(toJSON(msgData[pSlot]), "ahk_class AutoHotkey ahk_pid  " msgData[pSlot].pid)
}



; Read or write to ini file
; -------------------------
ini(job="read", var:="") 
{
	global
	local varsArry := isObject(var)? var : [var]
	
	if ( varsArry[1] == "" )
		return false

	for idx, varName in varsArry {
		if ( job == "read" ) {
			defaultVar := %varName%
			iniRead, %varName%, % mainAppName ".ini", Settings, % varName
			if ( !%varName% || %varName% == "ERROR" || %varName% == "" ) {
				%varName% := defaultVar
			}
		}
		else if ( job == "write" ) {
			if ( !%varName% || %varName% == "ERROR" || %varName% == "" )
				%varName% := %varName%
			iniWrite, % %varName%, % mainAppName ".ini", Settings, % varName
			;log("Saved " varName " with value " %varName%)
		}
	}
}



playSound() 
{
	SoundBeep, 300, 100
	SoundBeep, 600, 600
}



; Send data across script instances
; -------------------------------------------------------
sendAppMessage(ByRef StringToSend, ByRef TargetScriptTitle) 
{
  VarSetCapacity(CopyDataStruct, 3*A_PtrSize, 0)
  SizeInBytes := strPutVar(StringToSend, StringToSend, "utf-8")
  NumPut(SizeInBytes, CopyDataStruct, A_PtrSize)
  NumPut(&StringToSend, CopyDataStruct, 2*A_PtrSize)
  Prev_DetectHiddenWindows := A_DetectHiddenWindows
  Prev_TitleMatchMode := A_TitleMatchMode
  DetectHiddenWindows On
  SetTitleMatchMode 2
  SendMessage, 0x4a, 0, &CopyDataStruct,, % TargetScriptTitle
  DetectHiddenWindows %Prev_DetectHiddenWindows%
  SetTitleMatchMode %Prev_TitleMatchMode%
  return errorLevel
 }

strPutVar(string, ByRef var, encoding)
{
    varSetCapacity( var, StrPut(string, encoding) * ((encoding="utf-16"||encoding="cp1200") ? 2 : 1) )
    return StrPut(string, &var, encoding)
}


; runCMD v0.94 by SKAN on D34E/D37C @ autohotkey.com/boards/viewtopic.php?t=74647
; Based on StdOutToVar.ahk by Sean @ autohotkey.com/board/topic/15455-stdouttovar      

runCMD(CmdLine, workingDir:="", codepage:="CP0", Fn:="RunCMD_Output")  
{         
	local        			 		                                                       
	global a_Args					

	Fn := isFunc(Fn) ? func(Fn) : 0
	,dllCall("CreatePipe", "PtrP",hPipeR:=0, "PtrP",hPipeW:=0, "Ptr",0, "Int",0)
	,dllCall("SetHandleInformation", "Ptr",hPipeW, "Int",1, "Int",1)
	,dllCall("SetNamedPipeHandleState","Ptr",hPipeR, "UIntP",PIPE_NOWAIT:=1, "Ptr",0, "Ptr",0)
	,P8 := (A_PtrSize=8)
	,varSetCapacity(SI, P8 ? 104 : 68, 0)                          ; STARTUPINFO structure      
	,numPut(P8 ? 104 : 68, SI)                                     ; size of STARTUPINFO
	,numPut(STARTF_USESTDHANDLES:=0x100, SI, P8 ? 60 : 44,"UInt")  ; dwFlags
	,numPut(hPipeW, SI, P8 ? 88 : 60)                              ; hStdOutput
	,numPut(hPipeW, SI, P8 ? 96 : 64)                              ; hStdError
	,varSetCapacity(PI, P8 ? 24 : 16)                              ; PROCESS_INFORMATION structure

	if not dllCall("CreateProcess", "Ptr",0, "Str",CmdLine, "Ptr",0, "Int",0, "Int",True,"Int",0x08000000 | dllCall("GetPriorityClass", "Ptr",-1, "UInt"), "Int",0,"Ptr", workingDir ? &workingDir : 0, "Ptr",&SI, "Ptr",&PI)  
		return Format("{1:}", "", ErrorLevel := -1, dllCall("CloseHandle", "Ptr",hPipeW), dllCall("CloseHandle", "Ptr",hPipeR))

	dllCall("CloseHandle", "Ptr",hPipeW)
	,a_Args.runCMD := {"PID": NumGet(PI, P8 ? 16 : 8, "UInt")}
	,file := fileOpen(hPipeR, "h", codepage)
	,lineNum := 1,  sOutput := ""
	while ( a_Args.runCMD.PID + dllCall("Sleep", "Int", 50) && dllCall("PeekNamedPipe", "Ptr",hPipeR, "Ptr",0, "Int",0, "Ptr",0, "Ptr",0, "Ptr",0) ) {
		while ( a_Args.runCMD.PID && (line := file.readLine()) ) {
			sOutput .= Fn ? Fn.call(line, lineNum++, a_Args.runCMD.PID) : line
		}
	}
	
	a_Args.runCMD.PID := 0
	hProcess := numGet(PI, 0), hThread  := numGet(PI, a_PtrSize)
	,dllCall("GetExitCodeProcess", "Ptr",hProcess, "PtrP",ExitCode:=0), dllCall("CloseHandle", "Ptr",hProcess)
	,dllCall("CloseHandle", "Ptr",hThread), dllCall("CloseHandle", "Ptr",hPipeR)
	
	return {"msg":sOutput, "exitcode":ExitCode}
}


; Create a folder
; ---------------------------------------------
createFolder(newFolder) 
{
	if ( fileExist(newFolder) == "D" ) {	; Folder exists
		createdFolder := newFolder
	}
	else {
		if ( !splitPath(newFolder).drv ) {																							; No drive letter can be assertained, so it's invalid
			createdFolder := false
		} else { 																								; Output folder is valid but dosent exist
			fileCreateDir, % regExReplace(newFolder, "\\$")
			createdFolder := errorLevel? false : newFolder
		}
	}
	return createdFolder	; Returns the folder name if created or it exists, or false if no folder was created
}

; Kill all namDHC process (including chdman.exe)
; -----------------------------------------------
killAllProcess() 
{
	global mainAppName, mainAppNameVerbose, runAppName, runAppNameConsole

	while ( true ) {
		process, close, % "chdman.exe"
		if ( !errorLevel )
			break
	}

	for idx, app in [mainAppName, mainAppNameVerbose, runAppName, runAppNameConsole] {
		hwnd := winExist(app)
		winActivate % "ahk_id " hwnd
		winClose % "ahk_id " hwnd
		if ( winExist("ahk_id " hwnd) ) {
			postMessage, 0x0112, 0xF060,,, % "ahk_id " hwnd
			winKill % "ahk_id " hwnd
		}
	}
}

; Check if menu item has a checkmark (is checked)
; -----------------------------------------------
isMenuChecked(menuName, itemNumber)  
{
   static MIIM_STATE := 1, MFS_CHECKED := 0x8
   hMenu := MenuGetHandle(menuName)
   VarSetCapacity(MENUITEMINFO, size := 4*4 + A_PtrSize*8, 0)
   NumPut(size, MENUITEMINFO)
   NumPut(MIIM_STATE, MENUITEMINFO, 4, "UInt")
   DllCall("GetMenuItemInfo", Ptr, hMenu, UInt, itemNumber - 1, UInt, true, Ptr, &MENUITEMINFO)
   return !!(NumGet(MENUITEMINFO, 4*3, "UInt") & MFS_CHECKED)
}


; Disable a windows close button
; ------------------------------
disableCloseButton(hWnd="") 
{
	If ( hWnd == "" )
		hWnd := winExist("A")
	hSysMenu := dllCall("GetSystemMenu","Int",hWnd,"Int",FALSE)
	nCnt := dllCall("GetMenuItemCount","Int",hSysMenu)
	dllCall("RemoveMenu","Int",hSysMenu,"UInt",nCnt-1,"Uint","0x400")
	dllCall("RemoveMenu","Int",hSysMenu,"UInt",nCnt-2,"Uint","0x400")
	dllCall("DrawMenuBar","Int",hWnd)
	return ""
}


; Merge two objects or arrays
; -------------------------
mergeObj(sourceObj, targetObj) 
{
	targetObj := isObject(targetObj) ? targetObj : {}
	for k, v In sourceObj {
		if ( isObject(v) ) {
			if ( !targetObj.hasKey(k) )
				targetObj[k] := {}
			mergeObj(v, targetObj[k])
		} else
			targetObj[k] := v
	}
}

; Show an object as text
; ----------------------
showObj(obj, s := "") 
{
	static str
	if (s == "")
		str := ""
	for idx, v In obj {
		n := (s == "" ? idx : s . ", " . idx)
		if isObject(v)
			showObj(v, n)
		else
			str .= "[" . n . "] = " . v . "`r`n"
	}
	return rTrim(str, "`r`n")
}




procCountDDList()
{
	loop % envGet("NUMBER_OF_PROCESSORS")
		lst .= a_index "|"					; Create processor count dropdown list
	return "|" lst "|"						; Last "|" is to select last as default
}


splitPath(inputFile) 
{
	splitPath inputFile, file, dir, ext, noext, drv
	return {file:file, dir:dir, ext:ext, noext:noext, drv:drv}
}



LV_GetText2(row, byRef rtn="") 	; To allow an inline call of the LV_GetText() function
{
	rtn := LV_GetText(str, row)
	return str
}




; Get position of a GUI control
; --------------------------------
guiGetCtrlPos(ctrl, guiNum:=1) 
{
	guiControlGet, rtn, %guiNum%:Pos, %ctrl%
	return {x:rtnx, y:rtny, w:rtnw, h:rtnh}
}

; Convert string to lowercase
; ----------------------------
stringLower(str) 
{
	stringLower, rtn, str
	return rtn
}

;Convert string to uppercase
; ---------------------------
stringUpper(str, title:=false) 
{
	stringUpper, rtn, str, % title? "T":""
	return rtn
}

; Check if value is in array
; ---------------------------
inArray(cVal, thisArray) 
{
	for idx, val in thisArray 
		if ( cVal == val )
			return true
	return false
}


; Convert array to string of items seperated by commas
; -----------------------------------------------------
arrayToComma(thisArray, delim:=", ") 
{
	rtn := ""
	for idx, val in thisArray
		rtn .= val delim
	return regExReplace(rtn, delim, "", "", 1, -1)
}


; Remove an item from an array
; -------------------------------
removeFromArray(removeItem, byRef thisArray)
{
	if ( !isObject(thisArray) )
		return thisArray
	
	for idx, val in thisArray {
		if ( val == removeItem ) {
			thisArray.removeAt(idx)
			break
		}
	}
	return thisArray
}

; Get Windows current default font
; ------------------------
guiDefaultFont() ; by SKAN
{ 
   varSetCapacity(LF, szLF := 28 + (A_IsUnicode ? 64 : 32), 0) ; LOGFONT structure
   if DllCall("GetObject", "Ptr", DllCall("GetStockObject", "Int", 17, "Ptr"), "Int", szLF, "Ptr", &LF)
      return {name: StrGet(&LF + 28, 32), size: Round(Abs(NumGet(LF, 0, "Int")) * (72 / A_ScreenDPI), 1)
            , weight: NumGet(LF, 16, "Int"), quality: NumGet(LF, 26, "UChar")}
   return False
}


; GUI window was moved
; --------------------
moveGUIWin(wParam, lParam)
{
    global mainWinPosX, mainWinPosY, verboseWinPosX, verboseWinPosY, mainAppName, mainAppNameVerbose
	
	if ( a_gui == 1 ) {
		winGetPos, mainWinPosX, mainWinPosY,,, % mainAppName
		setTimer, writemoveGUIWin, -1000
	}
	else if ( a_gui == 2 ) {
		winGetPos, verboseWinPosX, verboseWinPosY,,, % mainAppNameVerbose
		setTimer, writemoveGUIWin, -1000
	}
	return   
	writemoveGUIWin:
		ini("write", ["mainWinPosX", "mainWinPosY"])
		ini("write", ["verboseWinPosX", "verboseWinPosY"])
	return
}


; Verbose window was resized
; --------------------------
2GuiSize(guiHwnd, eventInfo, W, H) 
{
	global verboseWinPosH := H, verboseWinPosW := W
	
	autoXYWH("wh", "editVerbose") 						; Resize edit control with window
	setTimer, write2GuiSize, -1000					
	return
	write2GuiSize:
		ini("write", ["verboseWinPosH", "verboseWinPosW"])
	return	
}


; =================================================================================
; Function: AutoXYWH
;   Move and resize control automatically when GUI resizes.
; Parameters:
;   DimSize - Can be one or more of x/y/w/h  optional followed by a fraction
;             add a '*' to DimSize to 'MoveDraw' the controls rather then just 'Move', this is recommended for Groupboxes
;   cList   - variadic list of ControlIDs
;             ControlID can be a control HWND, associated variable name, ClassNN or displayed text.
;             The later (displayed text) is possible but not recommend since not very reliable 
; Examples:
;   AutoXYWH("xy", "Btn1", "Btn2")
;   AutoXYWH("w0.5 h 0.75", hEdit, "displayed text", "vLabel", "Button1")
;   AutoXYWH("*w0.5 h 0.75", hGroupbox1, "GrbChoices")
; ---------------------------------------------------------------------------------
; Version: 2015-5-29 / Added 'reset' option (by tmplinshi)
;          2014-7-03 / toralf
;          2014-1-2  / tmplinshi
; requires AHK version : 1.1.13.01+
; =================================================================================
autoXYWH(DimSize, cList*)  ; http://ahkscript.org/boards/viewtopic.php?t=1079
{      
  static cInfo := {}
 
  If (DimSize = "reset")
    Return cInfo := {}
 
  for i, ctrl in cList {
    ctrlID := A_Gui ":" ctrl
    If ( cInfo[ctrlID].x = "" ){
        guiControlGet, i, %A_Gui%:Pos, %ctrl%
        MMD := InStr(DimSize, "*") ? "MoveDraw" : "Move"
        fx := fy := fw := fh := 0
        For i, dim in (a := StrSplit(RegExReplace(DimSize, "i)[^xywh]")))
            If !RegExMatch(DimSize, "i)" dim "\s*\K[\d.-]+", f%dim%)
              f%dim% := 1
        cInfo[ctrlID] := { x:ix, fx:fx, y:iy, fy:fy, w:iw, fw:fw, h:ih, fh:fh, gw:A_GuiWidth, gh:A_GuiHeight, a:a , m:MMD}
    }Else If ( cInfo[ctrlID].a.1) {
        dgx := dgw := A_GuiWidth  - cInfo[ctrlID].gw  , dgy := dgh := A_GuiHeight - cInfo[ctrlID].gh
        For i, dim in cInfo[ctrlID]["a"]
            Options .= dim (dg%dim% * cInfo[ctrlID]["f" dim] + cInfo[ctrlID][dim]) A_Space
        GuiControl, % A_Gui ":" cInfo[ctrlID].m , % ctrl, % Options
} } }


; Function replacement for guiControl
; ------------------------------------
; Example usages:
; guiCtrl({thisButton:"New Button Text"}, 1) - works on GUI #1
; guiCtrl("move", {stuff:"x9 w200", thing:"x1 y2"}, 3)

guiCtrl(arg1:="", arg2:="", arg3:="") 
{
	if ( isObject(arg1) )
		obj := arg1, guiNum := arg2 ? arg2 : 1
	else
		obj := arg2, cmd := arg1, guiNum := arg3 ? arg3 : 1

	for ele, newVal in obj
		guiControl, %guiNum%:%cmd%, % ele, % newVal
} 


; Function replacement for guiControlGet
; --------------------------------------
guiCtrlGet(ctrl, guiNum:=1) 
{
	guiControlGet, rtn, %guiNum%:, %ctrl%
	return rtn
}


; Draw spaces by count
; --------------------
drawSpace(num:=1) 
{
	if ( num < 1 ) 
		return ""
	loop % num
		rtn .= a_space
	return rtn
}

; Draw a line by count
; --------------------
drawLine(num:=1) 
{
	if ( num < 1 )
		return ""
	loop % num
		rtn .= "─"
	return rtn	
}

; Get an envirmoent variable
; --------------------------
envGet(enviro) 
{
	envGet, rtn, % enviro
	return rtn
}


; Delete a file 
; -------------
fileDelete(file, attempts:=5, sleepdelay:=200) 
{
	loop % (attempts < 1 ? 1 : attempts) { 			; 5 attempts to delete the file
		fileDelete, % file
		if ( errorLevel == 0 )						; Success
			return true
		sleep % sleepdelay
	}
	return false
}


; Delete a folder
; ---------------
folderDelete(dir, attempts:=5, sleepdelay:=200) 
{
	loop % (attempts < 1 ? 1 : attempts) {
		fileRemoveDir % dir, 0						; Attempt to delete the directory 5 times
		if ( errorLevel == 0 )						; Success
			return true
		sleep % sleepdelay
	}
	return false
}


; Attempt to close a process by PID
; ---------------------------------
processPIDClose(procPID, attempts:=5, sleepdelay:=200) 
{
	loop % attempts {
		process, close, % procPID
		if ( errorLevel == procPID )				; ErrorLevel returns PID if successful, or 0 if unsuccessful
			return true
		sleep % sleepdelay
	}
	return false
}


; Close App
; ---------
GuiClose()
{
	quitApp()
	exitApp			 ; Just in case
}

quitApp() 
{
	exitApp
}
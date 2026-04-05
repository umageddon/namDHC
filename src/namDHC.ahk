#singleInstance Off
DetectHiddenWindows(true)
SetTitleMatchMode(3)
APP_ROOT_DIR := A_ScriptDir
if ( !A_IsCompiled && InStr(FileExist(A_ScriptDir "\..\..\INPUT"), "D") )
	APP_ROOT_DIR := A_ScriptDir "\..\.."
else if ( !A_IsCompiled && InStr(FileExist(A_ScriptDir "\..\INPUT"), "D") )
	APP_ROOT_DIR := A_ScriptDir "\.."

if ( A_IsCompiled )
	chdmanCandidates := [APP_ROOT_DIR "\chdman.exe", APP_ROOT_DIR "\src\chdman.exe", APP_ROOT_DIR "\..\chdman.exe"]
else
	chdmanCandidates := [A_ScriptDir "\chdman.exe", APP_ROOT_DIR "\chdman.exe", A_ScriptDir "\..\chdman.exe"]

CHDMAN_FILE_LOC := ""
for _, thisPath in chdmanCandidates {
	if ( FileExist(thisPath) ) {
		CHDMAN_FILE_LOC := thisPath
		break
	}
}
if ( !CHDMAN_FILE_LOC )
	CHDMAN_FILE_LOC := APP_ROOT_DIR "\chdman.exe"
SetWorkingDir(APP_ROOT_DIR)


#Include SelectFolderEx.ahk
#Include ConsoleClass.ahk
#Include JSON.ahk


VER_HISTORY := "
(
v1.00
- Initial release


v1.01
- Added ISO input media support
- Minor fixes


v1.02
- Removed superfluous code
- Some spelling mistakes fixed
- Minor GUI changes


v1.03
- Fixed Cancel all jobs button
- Fixed output folder editfield allowing invalid characters
- Fixed files only being removed from listview after successful operation (when selected)
- Minor GUI bugs fixed
- GUI changes
- Added time elapsed to report
- Changed about window


v1.05
- Added update functionality
- GUI changes
- More GUI changes
- Fixed issue extracting or creating all formats instead of selected 
- Fixed JSON issues
- Bug fixes


v1.06
- Allows creation of multiple formats during a single job.
  Formats will be added to the file and directory names in parenthesis.
  Example of files from a PSX game with both TOC and CUE formats selected for output:
	PSXGAME (TOC).toc
	PSXGAME (TOC).bin
	PSXGAME (CUE).bin
	PSXGAME (CUE).cue
- User will be prompted to rename any duplicate output files
- Changed the finished job sound to play media instead of just beeps
- Fixed main menu not hiding during job processes
- Updater wont automatically update even when selecting no
- Bugfixes, GUI changes n' stuff


v1.07
- Added zip file support
  NOTE:  Folders & multiple formats within zip files are unsupported. 
         namDHC will only uncompress the first supported file that it finds within in each zipfile
- Changed quit routine
- Changed error handling for output chd files that already exist
- Changed file read functions
- Fixed namDHC won't ask about duplicate files when verifying or getting info from CHD's
- Fixed Folder and file browsing shows input extensions that aren't actually selected
- Fixed timeout monitoring
- Fixed some race conditions
- GUI changes n' stuff
- Changed JSON library again (hopefully last time)
- Having issues with script pausing/hanging after cancel command is sent to thread


v1.08
- Fixed Verify option
- No need to hit enter to confirm a new output folder
- Limit output folder name to 255 characters
- Removed Add/Remove metadata indefinitely


v1.09
- Fixed select output folder button not launching explorer window
- Fixed sometimes showing multiple jobs in one progress slot
- Re-enabled Autohotkey 'speedups'
- Fixed line character being misrepresented in report


v1.10
- Possible fix crashing and slowdown on some machines (Removed Autohotkey speedups :P)
- Fixed namDHC sometimes deleting the output file if it already existed
- Fixed a possible chdman timeout issue
- Fixed reporting to not show old reports from previous jobs


v1.11
- Possible crash fixes (?)
- Fixed zip files not being cleared from list after a successful job
- Slight speed improvement and (hopefully) improved reliability when cancelling jobs

v1.12
- Added an option to force no DPI scaling for people who are having issues with options being cut off in window (Thanks TFWol)
- Fixed quit routine

v1.13
- Final report shows size savings when compressing to CHD
- Fixed output folder being cleared if cancel was clicked in the select folder helper
- Removed close button on console windows
- Some cleanup

v2.0
- Migrated to AutoHotkey v2
- Fixes for GUI/threading/message handling, queue/cancel robustness and job execution stability
- Added 7-Zip support
- Added drag and drop support for input files. Drop folders or files on the main file list window to add
- Removed CHDMAN version checking
- Updated bundled CHDMAN to 0.287
- Added DVD create/extract support
- Removed batch updating - now we just go to github
- Other fixes and tons of changes
)"


; Default global values 
; ---------------------
CURRENT_VERSION := "2.0"
CHECK_FOR_UPDATES_STARTUP := "yes"
CHDMAN_FILE_LOC := CHDMAN_FILE_LOC ? CHDMAN_FILE_LOC : (APP_ROOT_DIR "\chdman.exe")
SEVENZIP_CUSTOM_PATH := ""
DIR_TEMP := a_Temp "\namDHC"
APP_MAIN_NAME := "namDHC"
APP_VERBOSE_NAME := APP_MAIN_NAME " - Verbose"
APP_RUN_JOB_NAME := APP_MAIN_NAME " - Job"
APP_RUN_CHDMAN_NAME := APP_RUN_JOB_NAME " - chdman"
APP_RUN_CONSOLE_NAME := APP_RUN_JOB_NAME " - Console"
RUNCMD_STATE := {PID:0}
APP_IS_CLOSING := false
gui6 := ""
APP_NO_DPI_SCALE := "no"
TIMEOUT_SEC := 20
WAIT_TIME_CONSOLE_SEC := 1
JOB_QUEUE_SIZE := 3
JOB_QUEUE_SIZE_LIMIT := 10
OUTPUT_FOLDER := a_workingDir
LAST_INPUT_BROWSE_FOLDER := InStr(FileExist(A_Desktop), "D") ? A_Desktop : a_workingDir
LAST_OUTPUT_BROWSE_FOLDER := InStr(FileExist(A_Desktop), "D") ? A_Desktop : a_workingDir
PLAY_SONG_FINISHED := "yes"
REMOVE_FILE_ENTRY_AFTER_FINISH := "yes"
SHOW_JOB_CONSOLE := "no"
SHOW_VERBOSE_WINDOW := "no"
USE_INPUTFOLDER_AS_OUTPUT := 0
APP_VERBOSE_WIN_HEIGHT := 400 
APP_VERBOSE_WIN_WIDTH := 800
APP_VERBOSE_WIN_POS_X := 775
APP_VERBOSE_WIN_POS_Y := 150
APP_MAIN_WIN_POS_X := 800
APP_MAIN_WIN_POS_Y := 100



; Read ini to write over globals if changed previously
;-------------------------------------------------------------
ini("read" 
	,["JOB_QUEUE_SIZE","OUTPUT_FOLDER","SHOW_JOB_CONSOLE","SHOW_VERBOSE_WINDOW","PLAY_SONG_FINISHED","REMOVE_FILE_ENTRY_AFTER_FINISH", "APP_NO_DPI_SCALE", "USE_INPUTFOLDER_AS_OUTPUT",
	"APP_MAIN_WIN_POS_X","APP_MAIN_WIN_POS_Y","APP_VERBOSE_WIN_WIDTH","APP_VERBOSE_WIN_HEIGHT","APP_VERBOSE_WIN_POS_X","APP_VERBOSE_WIN_POS_Y","CHECK_FOR_UPDATES_STARTUP","LAST_INPUT_BROWSE_FOLDER","LAST_OUTPUT_BROWSE_FOLDER","SEVENZIP_CUSTOM_PATH"])

if ( APP_MAIN_WIN_POS_X <= -32000 || APP_MAIN_WIN_POS_Y <= -32000 ) {
	APP_MAIN_WIN_POS_X := 800
	APP_MAIN_WIN_POS_Y := 100
	ini("write", ["APP_MAIN_WIN_POS_X", "APP_MAIN_WIN_POS_Y"])
}
if ( APP_VERBOSE_WIN_POS_X <= -32000 || APP_VERBOSE_WIN_POS_Y <= -32000 ) {
	APP_VERBOSE_WIN_POS_X := 775
	APP_VERBOSE_WIN_POS_Y := 150
	ini("write", ["APP_VERBOSE_WIN_POS_X", "APP_VERBOSE_WIN_POS_Y"])
}
savedBrowseDir := LAST_INPUT_BROWSE_FOLDER
LAST_INPUT_BROWSE_FOLDER := resolveNearestExistingDir(LAST_INPUT_BROWSE_FOLDER, InStr(FileExist(A_Desktop), "D") ? A_Desktop : a_workingDir)
if ( LAST_INPUT_BROWSE_FOLDER != savedBrowseDir )
	ini("write", "LAST_INPUT_BROWSE_FOLDER")
if ( !InStr(FileExist(LAST_INPUT_BROWSE_FOLDER), "D") ) {
	LAST_INPUT_BROWSE_FOLDER := InStr(FileExist(A_Desktop), "D") ? A_Desktop : a_workingDir
	ini("write", "LAST_INPUT_BROWSE_FOLDER")
}
savedOutputBrowseDir := LAST_OUTPUT_BROWSE_FOLDER
LAST_OUTPUT_BROWSE_FOLDER := resolveNearestExistingDir(LAST_OUTPUT_BROWSE_FOLDER, OUTPUT_FOLDER)
if ( LAST_OUTPUT_BROWSE_FOLDER != savedOutputBrowseDir )
	ini("write", "LAST_OUTPUT_BROWSE_FOLDER")
if ( !InStr(FileExist(LAST_OUTPUT_BROWSE_FOLDER), "D") ) {
	LAST_OUTPUT_BROWSE_FOLDER := resolveNearestExistingDir(OUTPUT_FOLDER, InStr(FileExist(A_Desktop), "D") ? A_Desktop : a_workingDir)
	ini("write", "LAST_OUTPUT_BROWSE_FOLDER")
}
if ( InStr(FileExist(OUTPUT_FOLDER), "D") && LAST_OUTPUT_BROWSE_FOLDER != OUTPUT_FOLDER ) {
	LAST_OUTPUT_BROWSE_FOLDER := OUTPUT_FOLDER
	ini("write", "LAST_OUTPUT_BROWSE_FOLDER")
}

if ( !fileExist(CHDMAN_FILE_LOC) ) {
	MsgBox("CHDMAN.EXE not found!`n`nMake sure the chdman executable is located in the same directory as namDHC and try again.", "Fatal Error", 16)
	ExitApp()
}

SEVENZIP_EXE := detect7ZipExe()
HAS_7ZIP := SEVENZIP_EXE != ""



; Run a chdman thread
; Will be called when running chdman - As to allow for a one file executable
;-------------------------------------------------------------
if ( A_Args.Length && A_Args[1] == "threadMode" ) {
	#include threads.ahk
}


; Kill all processes so only one instance is running
;-------------------------------------------------------------
killAllProcess()


; Set working job variables
;-------------------------------------------------------------
job := {started:false, workTally:{}, msgData:[], availPSlots:[], workQueue:[], scannedFiles:{}, queuedMsgData:[], InputExtTypes:[], OutputExtType:[], selectedOutputExtTypes:[], selectedInputExtTypes:[]}

; Set GUI variables
;-------------------------------------------------------------
UI := { chdmanOpt:{}, dropdowns:{job:{}}, media:{},  buttons:{normal:[], hover:[], clicked:[], disabled:[]}, menu:{namesOrder:[], File:[], Settings:[], About:[]}  }

; Legacy event state used to keep the current handler layout working under v2.
a_GuiControl := ""
a_GuiEvent := ""
a_ThisMenu := ""
a_ThisMenuItem := ""
a_ThisMenuItemPos := 0
gui1 := "", gui2 := "", gui3 := "", gui4 := "", gui5 := ""

UI.dropdowns.job := { create: {pos:1,desc:"Create CHD files from media"}, extract: {pos:2,desc:"Extract images from CHD files"}, info: {pos:3, desc:"Get info from CHD files"}, verify: {pos:4, desc:"Verify CHD files"} }
	/*
	,addMeta: {pos:5, desc:"Add metadata to CHD files"}
	,delMeta: {pos:6, desc:"Delete metadata from CHD files"} 
	*/
UI.dropdowns.media := { cd:"CD image", dvd:"DVD image", hd:"Hard disk image", ld:"LaserDisc image", raw:"Raw image" }

UI.buttons.default := { normal:[0, 0xFFCCCCCC, "", "", 3], hover:[0, 0xFFBBBBBB, "", 0xFF555555, 3], clicked:[0, 0xFFCFCFCF, "", 0xFFAAAAAA, 3], disabled:[0, 0xFFE0E0E0, "", 0xFFAAAAAA, 3] }
UI.buttons.cancel :=	{ normal:[0, 0xFFFC6D62, "", "White", 3], hover:[0, 0xFFff8e85, "", "White", 3], clicked:[0, 0xFFfad5d2, "", "White", 3], disabled:[0, 0xFFfad5d2, "", "White", 3]}
UI.buttons.start :=	{ normal:[0, 0xFF74b6cc, "", 0xFF444444, 3],	hover:[0, 0xFF84bed1, "", "White", 3], clicked:[0, 0xFFa5d6e6, "", "White", 3], disabled:[0, 0xFFd3dde0, "", 0xFF888888, 3] }	

; Set menu variables
;-------------------------------------------------------------
UI.menu.namesOrder := ["File", "Settings", "About"]
UI.menu.File.Push({name:"Quit",											gotolabel:"quitApp",					saveVar:""})
UI.menu.About.Push({name:"About",											gotolabel:"menuSelected",				saveVar:""})
UI.menu.Settings.Push({name:"Check for updates automatically",				gotolabel:"menuSelected",				saveVar:"CHECK_FOR_UPDATES_STARTUP"})
UI.menu.Settings.Push({name:"Number of jobs to run concurrently",			gotolabel:":SubSettingsConcurrently",	saveVar:""})
UI.menu.Settings.Push({name:"Show a verbose window",						gotolabel:"menuSelected",				saveVar:"SHOW_VERBOSE_WINDOW", 	Fn:"showVerbose"})
UI.menu.Settings.Push({name:"Show a chdman console for each new job",		gotolabel:"menuSelected",				saveVar:"SHOW_JOB_CONSOLE"})
UI.menu.Settings.Push({name:"Play a sound when finished all jobs",			gotolabel:"menuSelected",				saveVar:"PLAY_SONG_FINISHED"})
UI.menu.Settings.Push({name:"Remove files from list on success",			gotolabel:"menuSelected",				saveVar:"REMOVE_FILE_ENTRY_AFTER_FINISH"})
UI.menu.Settings.Push({name:"No DPI scaling",								gotolabel:"menuSelected",				saveVar:"APP_NO_DPI_SCALE", 	Fn: "msgNeedARestart"})
UI.menu.Settings.Push({name:"Set 7-Zip location...",						gotolabel:"set7ZipLocation",			saveVar:""})

; misc GUI variables
;-------------------------------------------------------------
UI.HDtemplate := { ddList: ""		; Hard drive template dropdown list
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
, values : [0,1,2,3,4,5,6,7,8,9,10,11,12]}
UI.CPUCores := procCountDDList()

/* 
GUI CHDMAN options

Format:
;-------------------------------------------------------------
	name:				String	- Friendly name of option - used as a reference
	paramString:		String	- String used in actual chdman command
	description:		String	- String used to describe option in the GUI
	editField:			String	- Creates an editfield for the option containting the string supplied
	  - useQuotes:		Boolean	- TRUE will add quotes around the users editfield when submitting to chdman - only used with editField
	dropdownOptions:	String	- Creates a dropdown with the options supplied.  Must be in autohotkey format (ie- seperated by '|')
	  - dropdownValues:	Array	- Used with dropdownOptions - Supplies a different set of values to submit when selecting a dropdown option
	hidden:				Boolean	- TRUE to hide the option in the GUI
	masterOf:			String	- Supply the element name this 'masters over' - the 'subordinate' option will only be enabled when this option is checked
	xInset:				Number	- Number to pixels move option right in the GUI
	checked:			Boolean - TRUE if checkbox is checked by default
	helpText:			String	- Optional concise hover text shown for the option controls
*/	

UI.chdmanOpt.force := { name: "force", paramString: "f", description: "Force overwriting an existing output file", isFileOption: true, helpText: "Overwrite an existing output file instead of skipping the job" }
UI.chdmanOpt.fixVerify := { name: "fixverify", paramString: "f", description: "Fix the SHA-1 if it is incorrect", helpText: "Update the CHD SHA-1 when verification finds a mismatch" }
UI.chdmanOpt.verbose := { name: "verbose", paramString: "v", description: "Verbose output", hidden: true, helpText: "Include more CHDMAN details in the report" }
UI.chdmanOpt.outputBin := { name: "outputbin", paramString: "ob",description: "Output filename for binary data",	editField: "filename.bin", useQuotes:true, helpText: "Set the BIN filename used when extracting CD media" }
UI.chdmanOpt.splitBin := { name: "splitbin", paramString: "sb", description: "Output one binary file per track", helpText: "Write one BIN file for each track instead of one combined BIN" }
UI.chdmanOpt.inputParent := { name: "inputparent", paramString: "ip", description: "Input Parent", editField: "filename.ext", useQuotes:true, helpText: "Use a parent CHD when reading a delta CHD" }
UI.chdmanOpt.inputStartFrame :=	{ name: "inputstartframe",paramString: "isf",description: "Input Start Frame",	editField: 0, helpText: "Start reading at this frame offset" }
UI.chdmanOpt.inputFrames := { name: "inputframes",paramString: "if",	description: "Effective length of input in frames",editField: 0, helpText: "Limit how many input frames are read" }
UI.chdmanOpt.inputStartByte := { name: "inputstartbyte",paramString: "isb",description: "Starting byte offset within the input", editField: 0, helpText: "Start reading at this byte offset" }
UI.chdmanOpt.outputParent :=	 {name: "outputparent",paramString: "op",description: "Parent CHD for output",editField: "parent.chd",useQuotes:true, helpText: "Create a delta CHD from this parent CHD without renaming the output file" }
UI.chdmanOpt.hunkSize :=	 {name: "hunksize",paramString: "hs",description: "Size of each hunk (in bytes)",editField: 19584, helpText: "Set the CHD hunk size in bytes for the output file" }
UI.chdmanOpt.inputStartHunk := {name: "inputstarthunk",paramString: "ish",description: "Starting hunk offset within the input",editField: 0, helpText: "Start reading at this hunk offset" }
UI.chdmanOpt.inputBytes := { name: "inputBytes",paramString: "ib",	description: "Effective length of input (in bytes)",editField: 0, helpText: "Limit how many input bytes are read" }
UI.chdmanOpt.compression := { name: "compression",paramString: "c",description: "Compression codecs to use",editField: "cdlz,cdzl,cdfl", helpText: "Choose the compression codecs CHDMAN should use" }
UI.chdmanOpt.inputHunks := { name: "inputhunks",paramString: "ih",	description: "Effective length of input (in hunks)",editField: 0, helpText: "Limit how many input hunks are read" }
UI.chdmanOpt.numProcessors := { name :"numprocessors",paramString: "np",	description: "Max number of CPU threads to use",dropdownOptions: UI.CPUCores, helpText: "Limit how many CPU threads CHDMAN may use" }
UI.chdmanOpt.template :=	 { name: "template",paramString: "tp",description: "Hard drive template to use",dropdownOptions: UI.HDtemplate.ddList,dropdownValues:UI.HDtemplate.values, helpText: "Use a preset hard disk layout instead of entering it manually" }
UI.chdmanOpt.chs := {
	name: "chs",
	paramString: "chs",
	description: "CHS Values [cyl, heads, sectors]",
	editField: "332,16,63",
	helpText: "Set hard disk geometry as cylinders, heads, sectors"
}
UI.chdmanOpt.ident := {
	name: "ident",
	paramString: "id",
	description: "Name of ident file for CHS info",
	editField: "filename.chs",
	useQuotes:true,
	helpText: "Load hard disk geometry from an ident file"
}
UI.chdmanOpt.size :=	 {
	name: "size",
	paramString: "s",
	description: "Size of output file (in bytes)",
	editField: 0,
	helpText: "Set the output size in bytes when creating a blank hard disk image"
}
UI.chdmanOpt.unitSize :=	 {
name: "unitsize",paramString: "us",description: "Size of each unit (in bytes)",editField: 0, helpText: "Set the addressable unit size in bytes for raw CHDs without a parent"
}
UI.chdmanOpt.sectorSize := {
name: "sectorsize",
paramString: "ss",
description: "Size of each hard disk sector (in bytes)",
editField: 512,
helpText: "Set the hard disk sector size in bytes"
}
UI.chdmanOpt.deleteInputFiles :={
name: "deleteInputFiles",
description: "Delete input files after completing job",
masterOf: "deleteInputDir",
isFileOption: true,
helpText: "Delete the source files after a successful job"
}
UI.chdmanOpt.deleteInputDir := {
name: "deleteInputDir",
description: "Also delete input directory (if empty)",
xInset:10,
isFileOption: true,
helpText: "Also remove the source folder if it is empty after deleting the source files"
}
UI.chdmanOpt.createSubDir :=	{
	name: "createSubDir",	description: "Create a folder for each successful job", isFileOption: true, helpText: "Save each job's output files in a separate folder"
}
UI.chdmanOpt.keepIncomplete :=		{name: "keepIncomplete",						description: "Keep failed or cancelled output files", isFileOption: true, helpText: "Keep partial output files when a job fails or is cancelled"}

setCurrentChdmanOptionDefaults(jobCmd:="")
{
	global UI

	UI.chdmanOpt.compression.editField := "cdlz,cdzl,cdfl"
	UI.chdmanOpt.hunkSize.editField := 19584
	UI.chdmanOpt.unitSize.editField := 512

	for optName in ["hunkSize", "unitSize", "createSubDir"] {
		if ( UI.chdmanOpt.%optName%.HasOwnProp("checked") )
			UI.chdmanOpt.%optName%.DeleteProp("checked")
	}

	switch jobCmd {
		case "createdvd", "createhd", "createraw":
			UI.chdmanOpt.compression.editField := "lzma,zlib,huff,flac"
			UI.chdmanOpt.hunkSize.editField := 4096

		case "createld":
			UI.chdmanOpt.compression.editField := "avhu"
			UI.chdmanOpt.hunkSize.editField := ""
	}

	if ( jobCmd == "createraw" ) {
		UI.chdmanOpt.hunkSize.checked := true
		UI.chdmanOpt.unitSize.checked := true
	}

	if ( RegExMatch(jobCmd, "^extract") )
		UI.chdmanOpt.createSubDir.checked := true
}


; Create Main GUI and its elements
;-------------------------------------------------------------
createMainGUI()
createProgressBars() 
createMenus()
gui1.Show('autosize x' APP_MAIN_WIN_POS_X ' y' APP_MAIN_WIN_POS_Y)
mainAppHWND := gui1.Hwnd

showVerbose(SHOW_VERBOSE_WINDOW)			; Check or uncheck item "Show verbose window"  and show the window 
selectJob()									; Select 1st selection in job dropdown list and trigger refreshGUI()

if ( CHECK_FOR_UPDATES_STARTUP == "yes" )
	checkForUpdates()

OnMessage(0x03,		moveGUIWin)			; If windows are moved, save positions in moveGUIWin()
OnMessage(0x200,	showCtrlHoverTip)

Sleep(25) 									; Needed (?) to allow window to be detected
if ( !mainAppHWND )
	mainAppHWND := winExist(APP_MAIN_NAME)

log(APP_MAIN_NAME " ready.")

return

;-------------------------------------------------------------------------------------------------------------------------


; ======================================================================================================================
; Menu handling and legacy event adapters
; ======================================================================================================================

; A Menu item was selected
;-------------------------------------------------------------
menuSelected() 
{
	global
	local selMenuObj, varName, fn
	
	switch a_ThisMenu {
		case "SettingsMenu":
			selMenuObj := UI.menu.settings[a_ThisMenuItemPos]									; Reference menu setting
			varName := selMenuObj.saveVar														; Get variable name
			currVal := namedSetting(varName)
			newVal := (currVal == "no") ? "yes" : "no"
			namedSetting(varName, newVal, true)												; Toggle variable setting
			if ( newVal == "yes" )
				UI.menuObj.SettingsMenu.Check(selMenuObj.name)
			else
				UI.menuObj.SettingsMenu.Uncheck(selMenuObj.name)
			ini("write", varName)																; Write new setting
			if ( selMenuObj.HasOwnProp("Fn") && selMenuObj.Fn != "" ) {														; Check if function needs to be called
				fn := selMenuObj.Fn
				switch fn {
					case "showVerbose":
						showVerbose(newVal)
					case "msgNeedARestart":
						msgNeedARestart()
				}
			}

		case "SubSettingsConcurrently":															; Menu: Settings: User selected number of jobs to run concurrently
			loop JOB_QUEUE_SIZE_LIMIT																; Uncheck all 
				UI.menuObj.SubSettingsConcurrently.Uncheck(a_index "&")
			UI.menuObj.SubSettingsConcurrently.Check(A_ThisMenuItemPos "&")					; Check selected
			JOB_QUEUE_SIZE := A_ThisMenuItemPos													; Set variable
			ini("write", "JOB_QUEUE_SIZE")
			log("Saved JOB_QUEUE_SIZE")
		
		case "AboutMenu":																		; Menu: About
			guiToggle("disable", "all")
				
			if ( gui4 )
				gui4.Destroy()
				
			gui4 := Gui("+OwnDialogs +LastFound +AlwaysOnTop +ToolWindow")
			gui4.OnEvent("Close", guiClose4)
			gui4.MarginX := 20 
			gui4.MarginY := 20
			gui4.SetFont('s20 Q5 w700 c000000')
			gui4.Add('text', 'x10 y10', APP_MAIN_NAME)
			gui4.setFont('s13 Q5 w700 c000000')
			gui4.Add('text', 'x+5 y17', " v" CURRENT_VERSION)
			gui4.SetFont('s10 Q5 w400 c000000')
			gui4.Add('text', 'x10 y40', "A Windows frontend for the MAME CHDMAN tool")
			ctrl := gui4.Add('button', 'x10 y70 w130 h22', "Check for updates")
			ctrl.OnEvent("Click", (*) => checkForUpdates())
			applyCtrlHoverTip(ctrl, "Check GitHub for a newer namDHC release")

			gui4.Add('text', 'x10 y110', "History")
			gui4.Add('edit', 'x10 y130 h200 w775', VER_HISTORY)
			gui4.Add('link', 'x10 y350', 'Github: <a href="https://github.com/umageddon/namDHC">https://github.com/umageddon/namDHC</a>')
			gui4.Add('link', 'x10 y370', 'MAME Info: <a href="https://www.mamedev.org/">https://www.mamedev.org/</a>')
			gui4.SetFont('s9 Q5 w400 c000000')
			gui4.Add('text', 'x10 y400', "(C) Copyright 2026 Umageddon")
			gui4.Show('w800 center')
			;controlFocus About 												; Removes outline around html anchor
			return
	}

}

guiClose4(*)
{
	global gui4
	gui4.Destroy()
	refreshGUI()
}


msgNeedARestart() {
	MsgBox("You will need to restart to see any changes", , 16)
}

set7ZipLocation()
{
	global gui1, SEVENZIP_CUSTOM_PATH, SEVENZIP_EXE, HAS_7ZIP, APP_MAIN_NAME, SB
	local startDir := "", selectedPath := "", resolvedExe := ""

	gui1.Opt("+OwnDialogs")
	startDir := get7ZipBrowseDir()
	selectedPath := FileSelect(1, startDir, "Select 7z.exe (look for the file named 7z.exe)", "7-Zip executable (7z.exe)")
	if ( !selectedPath )
		return

	resolvedExe := resolve7ZipPath(selectedPath)
	if ( !resolvedExe ) {
		MsgBox("The selected file is not 7z.exe.", APP_MAIN_NAME, 48)
		return
	}

	SEVENZIP_CUSTOM_PATH := selectedPath
	SEVENZIP_EXE := resolvedExe
	HAS_7ZIP := true
	ini("write", "SEVENZIP_CUSTOM_PATH")
	selectMedia()
	refreshGUI()
	log("Saved 7-Zip location: " SEVENZIP_EXE)
	SB.SetText("Saved 7-Zip location: " SEVENZIP_EXE, 1, 1)
}

legacyGuiEvent(handlerName, guiCtrlObj, eventInfo:="", *)
{
	global a_GuiControl, a_GuiEvent
	a_GuiControl := guiCtrlObj.Name
	a_GuiEvent := eventInfo
	if ( IsObject(handlerName) && HasMethod(handlerName, "Call") )
		return handlerName.Call()
	return
}

legacyGuiDropFiles(handlerName, guiObj, guiCtrlObj, fileArray, *)
{
	global a_GuiControl, a_GuiEvent
	a_GuiControl := IsObject(guiCtrlObj) ? guiCtrlObj.Name : ""
	a_GuiEvent := "DropFiles"
	if ( IsObject(handlerName) && HasMethod(handlerName, "Call") )
		return handlerName.Call(guiObj.Hwnd, fileArray)
	return
}

legacyMenuEvent(handlerName, menuName, itemName, itemPos, *)
{
	global a_ThisMenu, a_ThisMenuItem, a_ThisMenuItemPos
	a_ThisMenu := menuName
	a_ThisMenuItem := itemName
	a_ThisMenuItemPos := itemPos
	switch handlerName {
		case "menuSelected":
			return menuSelected()
		case "menuExtHandler":
			return menuExtHandler()
		case "set7ZipLocation":
			return set7ZipLocation()
		case "quitApp":
			return quitApp()
	}
	return
}

guiSubmit(guiObj)
{
	return guiObj.Submit(false)
}
; ======================================================================================================================
; Job and media selection
; ======================================================================================================================

; Job selection
;-------------------------------------------------------------
selectJob() 
{
	global
	local jobSel
	
	guiSubmit(gui1)
	jobSel := gui1["dropdownJob"].Text
	;gui1.Opt("+ownDialogs")
	
	; Changes depending on job selected
	switch jobSel {
		case UI.dropdowns.job.create.desc:
			newStartButtonLabel := "CREATE CHD"	
			guiCtrl({dropdownMedia:"|" UI.dropdowns.media.cd "|" UI.dropdowns.media.dvd "|" UI.dropdowns.media.hd "|" UI.dropdowns.media.ld "|" UI.dropdowns.media.raw})

		case UI.dropdowns.job.extract.desc:											
			newStartButtonLabel := "EXTRACT MEDIA"
			guiCtrl({dropdownMedia:"|" UI.dropdowns.media.cd "|" UI.dropdowns.media.dvd "|" UI.dropdowns.media.hd "|" UI.dropdowns.media.ld "|" UI.dropdowns.media.raw})

		case UI.dropdowns.job.info.desc:
			newStartButtonLabel := "GET INFO"
			guiCtrl({dropdownMedia:"|CHD Files"})

		case UI.dropdowns.job.verify.desc:
			guiToggle("enable", "all")
			newStartButtonLabel := "VERIFY CHD"
			guiCtrl({dropdownMedia:"|CHD Files"})
		
		case UI.dropdowns.job.addMeta.desc:
			newStartButtonLabel := "ADD METADATA"
			guiCtrl({dropdownMedia:"|CHD Files"})
			MsgBox("Option not implemented yet", , 64)
		
		case UI.dropdowns.job.delMeta.desc:
			newStartButtonLabel := "DELETE METADATA"
			guiCtrl({dropdownMedia:"|CHD Files"})
			MsgBox("Option not implemented yet", , 64)
	}
	
	guiCtrl({buttonStartJobs:newStartButtonLabel})																								; New start button label to reflect new job

	guiCtrl("choose", {dropdownMedia:"|1"}) 																										; Choose first item in media dropdown and fire the selection 
	selectMedia()
}


; Media selection
;-------------------------------------------------------------
selectMedia()
{
	global
	local jobSel, mediaSelText, mediaSel, key, opt, val, idx, optNum, checkOpt, changeOpt, changeOptVal, ctrlY, x, y, gPos, file, pos, maxBottom, baseGroupWidth, groupHeight
	local optPerSide:=9, ctrlH:=25
	
	guiSubmit(gui1)
	jobSel := gui1["dropdownJob"].Text
	mediaSelText := gui1["dropdownMedia"].Text
	;gui1.Opt("+ownDialogs")

	; User selected media
	switch mediaSelText {
		case UI.dropdowns.media.cd: 	mediaSel := "cd"
		case UI.dropdowns.media.dvd:	mediaSel := "dvd"
		case UI.dropdowns.media.hd:	mediaSel := "hd"
		case UI.dropdowns.media.ld:	mediaSel := "ld"
		case UI.dropdowns.media.raw:	mediaSel := "raw"
		default: mediaSel := "chd"
	}
	
	; Assign job variables according to media
	switch jobSel {
		case UI.dropdowns.job.create.desc:			job.Cmd := "create" mediaSel, 	job.Desc := "Create CHD from a " stringUpper(mediaSel) " image",	job.FinPreTxt := "Jobs created"
		case UI.dropdowns.job.extract.desc:		job.Cmd := "extract" mediaSel,	job.Desc := "Extract a " stringUpper(mediaSel) " image from CHD",	job.FinPreTxt := "Jobs extracted"
		case UI.dropdowns.job.info.desc:			job.Cmd := "info", 				job.Desc := "Get info from CHD",									job.FinPreTxt := "Read info from jobs"
		case UI.dropdowns.job.verify.desc:			job.Cmd := "verify",			job.Desc := "Verify CHD",											job.FinPreTxt := "Jobs verified"
		case UI.dropdowns.job.addMeta.desc:		job.Cmd := "addmeta", 			job.Desc := "Add Metadata to CHD",									job.FinPreTxt := "Jobs with metadata added to"
		case UI.dropdowns.job.delMeta.desc:		job.Cmd := "delmeta", 			job.Desc := "Delete Metadata from CHD",								job.FinPreTxt := "Jobs with metadata deleted from"
	}

	setCurrentChdmanOptionDefaults(job.Cmd)
	
	; Assign rest of job variables according to job
	switch job.Cmd {
		case "extractcd":	job.InputExtTypes := ["chd", "zip"],						job.OutputExtTypes := ["cue", "toc", "gdi"],	job.Options := [UI.chdmanOpt.force, UI.chdmanOpt.createSubDir, UI.chdmanOpt.deleteInputFiles, UI.chdmanOpt.keepIncomplete, UI.chdmanOpt.outputBin, UI.chdmanOpt.splitBin, UI.chdmanOpt.inputParent]
		case "extractdvd":	job.InputExtTypes := ["chd", "zip"],						job.OutputExtTypes := ["iso"],					job.Options := [UI.chdmanOpt.force, UI.chdmanOpt.createSubDir, UI.chdmanOpt.deleteInputFiles, UI.chdmanOpt.keepIncomplete, UI.chdmanOpt.inputParent, UI.chdmanOpt.inputStartByte, UI.chdmanOpt.inputStartHunk, UI.chdmanOpt.inputBytes, UI.chdmanOpt.inputHunks]
		case "extractld":	job.InputExtTypes := ["chd", "zip"],						job.OutputExtTypes := ["raw"],					job.Options := [UI.chdmanOpt.force, UI.chdmanOpt.createSubDir, UI.chdmanOpt.deleteInputFiles, UI.chdmanOpt.keepIncomplete, UI.chdmanOpt.inputParent, UI.chdmanOpt.inputStartFrame, UI.chdmanOpt.inputFrames]
		case "extracthd":	job.InputExtTypes := ["chd", "zip"],						job.OutputExtTypes := ["img"],					job.Options := [UI.chdmanOpt.force, UI.chdmanOpt.createSubDir, UI.chdmanOpt.deleteInputFiles, UI.chdmanOpt.keepIncomplete, UI.chdmanOpt.inputParent, UI.chdmanOpt.inputStartByte, UI.chdmanOpt.inputStartHunk, UI.chdmanOpt.inputBytes, UI.chdmanOpt.inputHunks]
		case "extractraw":	job.InputExtTypes := ["chd", "zip"],						job.OutputExtTypes := ["img", "raw"],			job.Options := [UI.chdmanOpt.force, UI.chdmanOpt.createSubDir, UI.chdmanOpt.deleteInputFiles, UI.chdmanOpt.keepIncomplete, UI.chdmanOpt.inputParent, UI.chdmanOpt.inputStartByte, UI.chdmanOpt.inputStartHunk, UI.chdmanOpt.inputBytes, UI.chdmanOpt.inputHunks]
		case "createcd": 	job.InputExtTypes := ["cue", "toc", "gdi", "iso", "zip"],	job.OutputExtTypes := ["chd"],					job.Options := [UI.chdmanOpt.force, UI.chdmanOpt.createSubDir, UI.chdmanOpt.deleteInputFiles, UI.chdmanOpt.deleteInputDir, UI.chdmanOpt.keepIncomplete, UI.chdmanOpt.numProcessors, UI.chdmanOpt.outputParent, UI.chdmanOpt.hunkSize, UI.chdmanOpt.compression]
		case "createdvd":	job.InputExtTypes := ["iso", "zip"],						job.OutputExtTypes := ["chd"],					job.Options := [UI.chdmanOpt.force, UI.chdmanOpt.createSubDir, UI.chdmanOpt.deleteInputFiles, UI.chdmanOpt.deleteInputDir, UI.chdmanOpt.keepIncomplete, UI.chdmanOpt.numProcessors, UI.chdmanOpt.outputParent, UI.chdmanOpt.inputStartByte, UI.chdmanOpt.inputStartHunk, UI.chdmanOpt.inputBytes, UI.chdmanOpt.inputHunks, UI.chdmanOpt.hunkSize, UI.chdmanOpt.compression]
		case "createld":	job.InputExtTypes := ["raw", "zip"],						job.OutputExtTypes := ["chd"],					job.Options := [UI.chdmanOpt.force, UI.chdmanOpt.createSubDir, UI.chdmanOpt.deleteInputFiles, UI.chdmanOpt.deleteInputDir, UI.chdmanOpt.keepIncomplete, UI.chdmanOpt.numProcessors, UI.chdmanOpt.outputParent, UI.chdmanOpt.inputStartFrame, UI.chdmanOpt.inputFrames, UI.chdmanOpt.hunkSize, UI.chdmanOpt.compression]
		case "createhd":	job.InputExtTypes := ["img", "zip"],						job.OutputExtTypes := ["chd"],					job.Options := [UI.chdmanOpt.force, UI.chdmanOpt.createSubDir, UI.chdmanOpt.deleteInputFiles, UI.chdmanOpt.deleteInputDir, UI.chdmanOpt.keepIncomplete, UI.chdmanOpt.numProcessors, UI.chdmanOpt.compression, UI.chdmanOpt.outputParent, UI.chdmanOpt.size, UI.chdmanOpt.inputStartByte, UI.chdmanOpt.inputStartHunk, UI.chdmanOpt.inputBytes, UI.chdmanOpt.inputHunks, UI.chdmanOpt.hunkSize, UI.chdmanOpt.ident, UI.chdmanOpt.template, UI.chdmanOpt.chs, UI.chdmanOpt.sectorSize]
		case "createraw":	job.InputExtTypes := ["img", "raw", "zip"],					job.OutputExtTypes := ["chd"],					job.Options := [UI.chdmanOpt.force, UI.chdmanOpt.createSubDir, UI.chdmanOpt.deleteInputFiles, UI.chdmanOpt.deleteInputDir, UI.chdmanOpt.keepIncomplete, UI.chdmanOpt.numProcessors, UI.chdmanOpt.outputParent, UI.chdmanOpt.inputStartByte, UI.chdmanOpt.inputStartHunk, UI.chdmanOpt.inputBytes, UI.chdmanOpt.inputHunks, UI.chdmanOpt.hunkSize, UI.chdmanOpt.unitSize, UI.chdmanOpt.compression]
		case "info":		job.InputExtTypes := ["chd", "zip"],						job.OutputExtTypes := [],						job.Options := []
		case "verify":		job.InputExtTypes := ["chd", "zip"],						job.OutputExtTypes := [],						job.Options := [UI.chdmanOpt.inputParent, UI.chdmanOpt.fixVerify]
		case "addmeta":		job.InputExtTypes := ["chd"],								job.OutputExtTypes := [],						job.Options := []
		case "delmeta":		job.InputExtTypes := ["chd"],								job.OutputExtTypes := [],						job.Options := []
	}	

	if ( HAS_7ZIP && inArray("zip", job.InputExtTypes) && !inArray("7z", job.InputExtTypes) )
		job.InputExtTypes.Push("7z")

	if ( !HAS_7ZIP && inArray("zip", job.InputExtTypes) )
		applyCtrlHoverTip(gui1["buttonInputExtType"], "Choose which input file types namDHC should include  7z is not installed on this system")
	else
		applyCtrlHoverTip(gui1["buttonInputExtType"], "Choose which input file types namDHC should include for this job")

	if ( !job.scannedFiles.HasOwnProp(job.Cmd) )
		job.scannedFiles.%job.Cmd% := []
	
	; Hide and uncheck ALL options
	for key, opt in UI.chdmanOpt.OwnProps() {
		guiToggle("hide", [opt.name "_checkbox", opt.name "_edit", opt.name "_dropdown"])	
		key := opt.name "_checkbox"
		guiCtrl({%key%:0})
	}
	
	; Show checkbox options depending on media selected
	if ( job.Options.Length )	{
		baseGroupWidth := guiGetCtrlPos("groupboxJob").w
		baseGroupY := guiGetCtrlPos("groupboxFileOptions").y
		nextGroupY := baseGroupY
		fileOpts := [], chdmanOpts := []

		for _, opt in job.Options {
			if ( !opt || (opt.HasOwnProp("hidden") && opt.hidden) )
				continue
			if ( opt.HasOwnProp("isFileOption") && opt.isFileOption )
				fileOpts.Push(opt)
			else
				chdmanOpts.Push(opt)
		}

		groupDefs := [
			{name: "groupboxFileOptions", options: fileOpts},
			{name: "groupboxOptions", options: chdmanOpts}
		]

		for _, groupDef in groupDefs {
			groupName := groupDef.name
			groupOpts := groupDef.options

			if ( !groupOpts.Length ) {
				guiToggle("hide", groupName)
				guiCtrl("moveDraw", {%groupName%:"x5 y" nextGroupY " w" baseGroupWidth " h0"})
				continue
			}

			guiCtrl("moveDraw", {%groupName%:"x5 y" nextGroupY " w" baseGroupWidth " h1"})
			guiToggle("show", groupName)
			gPos := guiGetCtrlPos(groupName)
			idx := 0, ctrlY := 0, x := gPos.x + 10, y := gPos.y
			maxBottom := gPos.y
			visibleOptCount := 0

			for _, opt in groupOpts {
				if ( idx >= optPerSide )
					x += 400, y := gPos.y, idx := 0
				idx++, ctrlY := idx
				visibleOptCount++

				checkOpt := opt.name "_checkbox"
				changeOpt := ""
				guiCtrl({%checkOpt%:" " (opt.description ? opt.description : opt.name)})
				guiCtrl("moveDraw", {%checkOpt%:"x" x + (opt.HasOwnProp("xInset") ? opt.xInset : 0) " y" y + (ctrlY * ctrlH)})
				pos := guiGetCtrlPos(checkOpt)
				maxBottom := Max(maxBottom, pos.y + pos.h)

				if ( opt.HasOwnProp("editField") ) {
					changeOpt := opt.name "_edit"
					guiCtrl({%changeOpt%:opt.editField})
				}
				else if ( opt.HasOwnProp("dropdownOptions") ) {
					changeOpt := opt.name "_dropdown"
					guiCtrl({%changeOpt%:opt.dropdownOptions})
				}

				if ( changeOpt ) {
					guiCtrl("moveDraw", {%changeOpt%:"x" x + 210 " y" y + (ctrlY * ctrlH) - 3})
					guiToggle("show", [checkOpt, changeOpt])
					pos := guiGetCtrlPos(changeOpt)
					maxBottom := Max(maxBottom, pos.y + pos.h)
				}
				else
					guiToggle("show", checkOpt)

				guiCtrl({%checkOpt%:(opt.HasOwnProp("checked") && opt.checked) ? 1 : 0})
			}

			groupHeight := Max(Ceil((visibleOptCount > optPerSide ? optPerSide : visibleOptCount) * ctrlH) + 30, (maxBottom - gPos.y) + 15)
			guiCtrl("moveDraw", {%groupName%:"x5 y" nextGroupY " w" baseGroupWidth " h" groupHeight})
			nextGroupY += groupHeight + (groupName == "groupboxFileOptions" ? 0 : 16)
		}
	}
	else {
		baseGroupWidth := guiGetCtrlPos("groupboxJob").w
		baseGroupY := guiGetCtrlPos("groupboxFileOptions").y
		guiToggle("hide", ["groupboxFileOptions", "groupboxOptions"])
		guiCtrl("moveDraw", {groupboxFileOptions:"x5 y" baseGroupY " w" baseGroupWidth " h0"})
		guiCtrl("moveDraw", {groupboxOptions:"x5 y" baseGroupY " w" baseGroupWidth " h0"})
	}
	
	; Reset extension menus
	menuExtHandler(true) 																

	; Create and populate listview
	if ( !LV )
		LV := gui1["listViewInputFiles"]
	LV.Delete()																			; Delete all listview entries
	for idx, file in job.scannedFiles.%job.Cmd%
		LV.Add("", file)																				; Re-populate listview with scanned files
	try ControlFocus("SysListView321", "ahk_id " gui1.Hwnd)														; Focus on listview to stop one item being selected
	
	refreshGUI()
}


; Refreshes GUI to reflect current settings
;-------------------------------------------------------------
refreshGUI() 
{
	global
	local jobSel, opt, key, optNum
	static selectedJob
	
	; Timers and late GUI events can still fire while a job run is active.
	; Do not rebuild the main layout during active work or the progress section gets reset/hidden.
	if ( job.HasOwnProp("started") && job.started )
		return

	guiSubmit(gui1)
	jobSel := gui1["dropdownJob"].Text
	
	if ( !LV )
		LV := gui1["listViewInputFiles"]
	LV.Redraw()
	
	; By default, enable all elements
	guiToggle("enable", "all")
	
	; Show the main menu
	toggleMainMenu("show")
	
	; Changes to elements depending on job selected
	switch jobSel {
		case UI.dropdowns.job.create.desc:
			
		case UI.dropdowns.job.extract.desc:											
			
		case UI.dropdowns.job.info.desc:
			guiToggle("disable", ["dropdownMedia", "buttonOutputExtType", "buttonInputExtType", "editOutputFolder", "buttonBrowseOutput"])

		case UI.dropdowns.job.verify.desc:
			guiToggle("disable", ["dropdownMedia", "buttonOutputExtType", "buttonInputExtType", "editOutputFolder", "buttonBrowseOutput", "buttonBrowseOutput"])
			
		case UI.dropdowns.job.addMeta.desc:
			guiToggle("disable", "all")
			guiToggle("enable", "dropdownJob")
		
		case UI.dropdowns.job.delMeta.desc:
			guiToggle("disable", "all")
			guiToggle("enable", "dropdownJob")
	}
	
	; Enable chdman option checkboxes depending on job selected
	for optNum, opt in job.Options															
		guiToggle("enable", opt.name "_checkbox")
	

	; Checked option: enable or disable chdman option editfields, dropdowns or slave options 
	for optNum, opt in job.Options														
	{
		toggleCtrls := [opt.name "_dropdown", opt.name "_edit"]
		if ( opt.HasOwnProp("masterOf") )
			toggleCtrls.Push(opt.masterOf "_checkbox", opt.masterOf "_dropdown", opt.masterOf "_edit")
		guiToggle((guiCtrlGet(opt.name "_checkbox") ? "enable":"disable"), toggleCtrls) ; If checked, enable or disable dropdown or editfields
	}

	
	; Hide progress bars, progress text & progress groupbox
	guiToggle("hide", ["progressAll", "progressTextAll", "groupBoxProgress"])
	loop JOB_QUEUE_SIZE_LIMIT {
		progressSetBusy(a_index, false)
		guiToggle("hide", ["progress" a_index, "progressText" a_index, "progressCancelButton" a_index])
	}


	; Trigger listview() to refresh selection buttons
	listViewInputFiles()
	
	; Set status text
	SB.SetText(job.scannedFiles.%job.Cmd%.Length? job.scannedFiles.%job.Cmd%.Length " jobs in the " stringUpper(job.Cmd) " queue. Ready to start!" : "Add files to the job queue to start", 1)
	
	; Make sure main button is showing
	guiToggle("hide", "buttonCancelAllJobs")
	guiToggle("show", "buttonStartJobs")
	guiToggle((job.scannedFiles.%job.Cmd%.Length>0 ? "enable":"disable"), ["buttonStartJobs", "buttonselectAllInputFiles"]) 		; Enable start button if there are jobs in the listview
	
	; Disable output folder edit field and select a folder button if the USE_INPUTFOLDER_AS_OUTPUT checkbox is checked
	guiToggle(USE_INPUTFOLDER_AS_OUTPUT ? "disable":"enable", ["buttonBrowseOutput", "editOutputFolder"])
	
	; Show and resize main GUI
	gui1.Show('autosize x' APP_MAIN_WIN_POS_X ' y' APP_MAIN_WIN_POS_Y)
	DllCall("RedrawWindow", "Ptr", gui1.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x585)
}


; ======================================================================================================================
; Input/output type menus and file list management
; ======================================================================================================================

; User pressed input or output files button
; Show Ext menu
;-------------------------------------------------------------
buttonExtSelect()
{
	global UI, a_GuiControl

	switch a_guicontrol {
		case "buttonInputExtType":
			showMenuUnderCtrl(UI.menuObj.InputExtTypes, "buttonInputExtType")
		case "buttonOutputExtType":
			showMenuUnderCtrl(UI.menuObj.OutputExtTypes, "buttonOutputExtType")
	}
}

	
; User selected an extension from the input/output extension menu
;-------------------------------------------------------------
menuExtHandler(init:=false)
{
	global job, UI, gui1, a_ThisMenu, a_ThisMenuItem, HAS_7ZIP
	local jobSel := gui1["dropdownJob"].Text, removedCount := 0
	
	if ( init == true ) {													; Create and populate extension menu lists
		for idx, type in ["InputExtTypes", "OutputExtTypes"] {	
			UI.menuObj.%type%.Delete()									; Clear all old Input & Output menu items
			
			job.%"selected" type% := []										; Clear global Array of selected Input & Output extensions
			for idx2, ext in job.%type% {									; Parse through job.InputExtTypes & job.OutputExtTypes
				if ( !ext )
					continue
				UI.menuObj.%type%.Add(ext, legacyMenuEvent.Bind("menuExtHandler", type))				; Add extension item to the menu
				if ( jobSel == UI.dropdowns.job.extract.desc && type == "OutputExtTypes" && idx2 > 1 )
					continue												; By default, only check one extension of the Output menu if we are extracting an image
				else {
					UI.menuObj.%type%.Check(ext)							; Otherwise, check all extension menu items
					job.%"selected" type%.push(ext)							; Then add it to the input & output global selected extension array
				}
			}
			if ( type == "InputExtTypes" && !HAS_7ZIP && inArray("zip", job.InputExtTypes) ) {
				UI.menuObj.%type%.Add("7z (not installed on this system)", legacyMenuEvent.Bind("menuExtHandler", type))
				UI.menuObj.%type%.Disable("7z (not installed on this system)")
			}
		}
	}
	else if ( a_ThisMenu ) {												; An extension menu was selected
		selectedExtList := "selected" strReplace(a_ThisMenu, "extTypes", "") "ExtTypes"
		job.%selectedExtList% := []											; Re-build either of these Arrays: job.selectedOutputExtTypes[] or job.selectedInputExtTypes[]
		
		switch a_ThisMenu {													; a_ThisMenu is either 'InputExtTypes' or 'OutputExtTypes'
			case "OutputExtTypes":											; Only one output extension is allowed to be checked
				for _, val in job.OutputExtTypes
					UI.menuObj.OutputExtTypes.Uncheck(val)
				UI.menuObj.OutputExtTypes.Check(a_ThisMenuItem)
		
			case "InputExtTypes": 
				UI.menuObj.InputExtTypes.ToggleCheck(a_ThisMenuItem)		; Toggle checking item
				
		}
		for idx, val in job.%a_ThisMenu%
			if ( isMenuChecked(a_ThisMenu, idx) ) {
				job.%selectedExtList%.push(val)								; Add checked extension item(s) to the global array for reference later
			}
		if ( job.%selectedExtList%.Length == 0 ) {
			UI.menuObj.%a_ThisMenu%.Check(a_ThisMenuItem)					; Make sure at least one item is checked
			job.%selectedExtList%.push(a_ThisMenuItem)
		}
		if ( a_ThisMenu == "InputExtTypes" )
			removedCount := pruneQueuedFilesBySelectedInputTypes()
	}
	
}




; Drag and drop files on listview to add
; ------------------------------------------
guiDropFiles(guiHWND, fileArray) {
	global a_GuiControl, SB
	
	if ( a_guiControl == "listViewInputFiles" ) {
		arr := []
		scanCount := 0
		SB.SetText("Scanning dropped items...", 1, 1)
		for idx, file in fileArray 
		{
			if ( fileExist(file) == "D" ) {
				loop Files file "/*.*", "FR" {
					arr.push(A_LoopFileFullPath)
					scanCount++
					if ( Mod(scanCount, 250) = 0 ) {
						SB.SetText("Scanning dropped folders... " scanCount " file(s) found", 1, 1)
						Sleep(-1)
					}
				}
			} else
				arr.push(file)
		}
		SB.SetText("Finished scanning dropped items. Adding files...", 1, 1)

		addFolderFiles(arr)
	}
}




; Scan files and add to queue
;-------------------------------------------------------------
addFolderFiles(newFiles:=false, ctrlName:="")
{
	global job, APP_MAIN_NAME, SB, gui1, LV, mainAppHWND, a_GuiControl, LAST_INPUT_BROWSE_FOLDER
	newFiles := isObject(newFiles) ? newFiles : [], extList := "", numAdded := 0, thisCtrl := ctrlName ? ctrlName : a_GuiControl, selectedInputExtTypesLower := []
	selectedZipSearchExtTypes := []
	selectedZipSearchExtTypesLower := []
	badZipSelections := []
	selectedInputTypeText := ""
	scanCount := 0
	
	guiSubmit(gui1)
	gui1.Opt('+OwnDialogs')

	if ( !job.scannedFiles.HasOwnProp(job.Cmd) ) ; Create a job filelist object if it dosent exist
		job.scannedFiles.%job.Cmd% := []

	for idx, thisExt in job.selectedInputExtTypes {
		selectedInputExtTypesLower.Push(StrLower(thisExt))
		if ( !isArchiveExt(thisExt) ) {
			selectedZipSearchExtTypes.Push(thisExt)
			selectedZipSearchExtTypesLower.Push(StrLower(thisExt))
		}
	}
	selectedInputTypeText := arrayToOrString(selectedZipSearchExtTypes)
	
	switch ( thisCtrl ) {
		case "buttonAddFiles":
			for idx, ext in job.selectedInputExtTypes
				extList .= "*." ext "; "
			extList := RTrim(extList, "; ")
			startDir := resolveNearestExistingDir(LAST_INPUT_BROWSE_FOLDER, InStr(FileExist(A_Desktop), "D") ? A_Desktop : "")
			if ( startDir != LAST_INPUT_BROWSE_FOLDER ) {
				LAST_INPUT_BROWSE_FOLDER := startDir
				ini("write", "LAST_INPUT_BROWSE_FOLDER")
			}
			clearCtrlHoverTip()
			newInputList := FileSelect("M", startDir, "Select files", extList ? "Supported files (" extList ")" : "")
			if ( isObject(newInputList) ) {
				for idx, selectedFile in newInputList
					newFiles.push(selectedFile)
				if ( newInputList.Length ) {
					selPath := newInputList[1]
					LAST_INPUT_BROWSE_FOLDER := InStr(FileExist(selPath), "D") ? selPath : splitFilePath(selPath).dir
					ini("write", "LAST_INPUT_BROWSE_FOLDER")
				}
			}
			else if ( newInputList ) {
				newFiles.push(newInputList)
				LAST_INPUT_BROWSE_FOLDER := InStr(FileExist(newInputList), "D") ? newInputList : splitFilePath(newInputList).dir
				ini("write", "LAST_INPUT_BROWSE_FOLDER")
			}
		
		case "buttonAddFolder":
			startDir := resolveNearestExistingDir(LAST_INPUT_BROWSE_FOLDER, InStr(FileExist(A_Desktop), "D") ? A_Desktop : "")
			if ( startDir != LAST_INPUT_BROWSE_FOLDER ) {
				LAST_INPUT_BROWSE_FOLDER := startDir
				ini("write", "LAST_INPUT_BROWSE_FOLDER")
			}
			clearCtrlHoverTip()
			inputFolderResult := SelectFolderEx(startDir, "Select a folder containing " arrayToString(job.selectedInputExtTypes) " type files.", mainAppHWND)
			inputFolder := IsObject(inputFolderResult) ? inputFolderResult.SelectedDir : inputFolderResult
			if ( inputFolder ) {
				inputFolder := regExReplace(inputFolder, "\\$")
				LAST_INPUT_BROWSE_FOLDER := inputFolder
				ini("write", "LAST_INPUT_BROWSE_FOLDER")
				SB.SetText("Scanning folder '" inputFolder "'...", 1, 1)
				
				for idx, thisExt in job.selectedInputExtTypes {
					loop Files inputFolder "\*." thisExt, "FR" {
						newFiles.push(A_LoopFileFullPath)
						scanCount++
						if ( Mod(scanCount, 250) = 0 ) {
							SB.SetText("Scanning folder... " scanCount " file(s) found", 1, 1)
							Sleep(-1)
						}
					}
				}
				SB.SetText("Finished scanning folder. Adding files...", 1, 1)
		}
		
	}
	
	if ( newFiles.Length ) {
		guiToggle("disable", "all")
		for idx, thisFile in newFiles {
			addFile := true, msg := "", matchedZipFile := "", fileParts := splitFilePath(thisFile), inputExtLower := StrLower(fileParts.ext)
			if ( Mod(idx, 100) = 0 ) {
				SB.SetText("Adding files... " idx "/" newFiles.Length, 1, 1)
				Sleep(-1)
			}
			
			if ( !inArray(inputExtLower, selectedInputExtTypesLower) )
				addFile := false ;,msg := "Skip adding " thisFile "  -  Not a selected format"

			else if ( inArray(thisFile, job.scannedFiles.%job.Cmd%) )
				addFile := false, msg := "Skip adding " thisFile "  -  Already in queue"

			else if ( isArchiveExt(inputExtLower) ) { 																; If its an archive file, check to see if user extensions are contained within it
				addFile := false, msg := "Skip adding " thisFile "  -  No selected formats found in archive file"
				for idx, fileInZip in readArchiveFile(thisFile) {
					zipFileExtLower := StrLower(getPathExt(fileInZip))
					if ( inArray(zipFileExtLower, selectedZipSearchExtTypesLower) ) {
						addFile := true
						matchedZipFile := fileInZip
						break
					}
				}
				if ( !addFile && thisCtrl == "buttonAddFiles" )
					badZipSelections.Push(thisFile)
			}

			if ( addFile ) {
				numAdded++
				msg := "Adding " (matchedZipFile ? thisFile "   -->   " matchedZipFile : thisFile)
				LV.Add("", thisFile)
				job.scannedFiles.%job.Cmd%.push(thisFile)
			}
			
			log(msg)
			SB.SetText(msg, 1, 1)
		}
		if ( badZipSelections.Length ) {
			displayTypeText := selectedInputTypeText ? selectedInputTypeText : "supported"
			msg := badZipSelections.Length == 1
				? "No " displayTypeText " media was found in the archive file:`n`n" badZipSelections[1]
				: "No " displayTypeText " media was found in these archive files:`n`n" arrayToString(badZipSelections, "`n")
			MsgBox(msg, APP_MAIN_NAME, "Icon!")
		}
	}
	reportQueuedFiles()
	refreshGUI()
}


; The listview containting the current input files was clicked
;-------------------------------------------------------------
listViewInputFiles()
{
	global
	local suffx, idx, val
	
	if ( LV.GetCount("S") > 0 )  {
		suffx := (LV.GetCount("S") > 1) ? "s" : ""
		guiCtrl({buttonRemoveInputFiles:"Remove file" suffx, buttonclearInputFiles:"Unselect"})
		guiToggle("enable", ["buttonRemoveInputFiles", "buttonclearInputFiles"])
	}
	else
		guiToggle("disable", ["buttonclearInputFiles", "buttonRemoveInputFiles"]) 
		
	guiToggle((LV.GetCount()>0 ? "enable":"disable"), ["buttonselectAllInputFiles"]) 	
}


; Select file(s) from current input listview 
; ------------------------------------------
selectInputFiles()
{
	global job, LV, gui1, SB, a_GuiControl
	row := 0, removeThese := []
	
	guiSubmit(gui1)
	gui1.Opt("+OwnDialogs")
	
	if ( inStr(a_GuiControl, "SelectAll") ) {
		LV.Modify(0, "Select")
	}
	else if ( inStr(a_GuiControl, "Clear") ) {
		LV.Modify(0, "-Select")
	}
	else if ( inStr(a_GuiControl, "Remove") ) {
		guiToggle("disable", "all")
		loop {
			row := LV.GetNext(row)										; Get selected download from list and move to next
			if ( !row ) 												; Break if no more selected
				break
			removeThese.push(row)
		}
		while ( removeThese.Length ) {
			row := removeThese.pop()
			removeThisFile := LV.GetText(row, 1)
			LV.Delete(row)
			removeFromArray(removeThisFile, job.scannedFiles.%job.Cmd%)
			log("Removed '" removeThisFile "' from the " stringUpper(job.Cmd) " queue")
			SB.SetText("Removed '" removeThisFile "' from the " stringUpper(job.Cmd) " queue", 1)
		}
		reportQueuedFiles()
	}
	refreshGUI()
	try ControlFocus("SysListView321", "ahk_id " gui1.Hwnd)
}


; Report Queued files to the bottom info bar
; ------------------------------------------
reportQueuedFiles() 
{
	global job, SB
	log( job.scannedFiles.%job.Cmd%.Length? job.scannedFiles.%job.Cmd%.Length " jobs in the " stringUpper(job.Cmd) " queue. Ready to start!" : "No jobs in the " stringUpper(job.Cmd) " file queue" )
	SB.SetText( job.scannedFiles.%job.Cmd%.Length? job.scannedFiles.%job.Cmd%.Length " jobs in the " stringUpper(job.Cmd) " queue. Ready to start!" : "No jobs in the " stringUpper(job.Cmd) " file queue" )
}

pruneQueuedFilesBySelectedInputTypes()
{
	global job, LV, gui1
	local selectedInputExtTypesLower := [], keptFiles := [], removedCount := 0, thisFile

	if ( !job.scannedFiles.HasOwnProp(job.Cmd) )
		return 0

	for _, thisExt in job.selectedInputExtTypes
		selectedInputExtTypesLower.Push(StrLower(thisExt))

	for _, thisFile in job.scannedFiles.%job.Cmd% {
		if ( queuedFileMatchesSelectedInputTypes(thisFile, selectedInputExtTypesLower) )
			keptFiles.Push(thisFile)
		else
			removedCount++
	}

	if ( removedCount < 1 )
		return 0

	job.scannedFiles.%job.Cmd% := keptFiles

	if ( !LV )
		LV := gui1["listViewInputFiles"]
	LV.Delete()
	for _, thisFile in keptFiles
		LV.Add("", thisFile)

	try ControlFocus("SysListView321", "ahk_id " gui1.Hwnd)
	listViewInputFiles()
	log("Removed " removedCount " queued file(s) that no longer match the selected input file types")
	reportQueuedFiles()
	return removedCount
}

queuedFileMatchesSelectedInputTypes(thisFile, selectedInputExtTypesLower)
{
	local inputExtLower := StrLower(splitFilePath(thisFile).ext), zipFileExtLower

	if ( !inArray(inputExtLower, selectedInputExtTypesLower) )
		return false

	if ( !isArchiveExt(inputExtLower) )
		return true

	for _, fileInZip in readArchiveFile(thisFile) {
		zipFileExtLower := StrLower(getPathExt(fileInZip))
		if ( inArray(zipFileExtLower, selectedInputExtTypesLower) && !isArchiveExt(zipFileExtLower) )
			return true
	}
	return false
}
; ======================================================================================================================
; Output folder and option control handling
; ======================================================================================================================

; Select output folder
; --------------------
editOutputFolder(ctrlName:="")
{
	global a_GuiControl
	local thisCtrl := ctrlName ? ctrlName : a_GuiControl
	SetTimer(() => checkNewOutputFolder(thisCtrl), -500)
	return
}


; Check new inputted folder
; -------------------------
checkNewOutputFolder(ctrlName:="")
{
	global gui1, OUTPUT_FOLDER, LAST_OUTPUT_BROWSE_FOLDER, SB, mainAppHWND, a_GuiControl
	local thisCtrl := ctrlName ? ctrlName : a_GuiControl
	guiSubmit(gui1)
	gui1.Opt("+ownDialogs")

	if ( thisCtrl == "buttonBrowseOutput" ) {
		currentFolder := Trim(gui1["editOutputFolder"].Value)
		LAST_OUTPUT_BROWSE_FOLDER := resolveNearestExistingDir(LAST_OUTPUT_BROWSE_FOLDER, OUTPUT_FOLDER)
		ini("write", "LAST_OUTPUT_BROWSE_FOLDER")
		startDir := InStr(FileExist(currentFolder), "D")
			? currentFolder
			: LAST_OUTPUT_BROWSE_FOLDER
		clearCtrlHoverTip()
		selectFolderResult := SelectFolderEx(startDir, "Select a folder to save converted files to", mainAppHWND)
		selectFolder := IsObject(selectFolderResult) ? selectFolderResult.SelectedDir : selectFolderResult
		if ( selectFolder ) {
			LAST_OUTPUT_BROWSE_FOLDER := resolveNearestExistingDir(selectFolder, startDir)
			ini("write", "LAST_OUTPUT_BROWSE_FOLDER")
			guiCtrl({editOutputFolder:selectFolder}) ; Assign to edit field if user selected
		}
	}

	newFolder := gui1["editOutputFolder"].Value
	
	if ( !newFolder || newFolder == OUTPUT_FOLDER ) {
		newFolder := OUTPUT_FOLDER 								; Return to old value if no change or no value given
		return
	}

	badChar := false 											
	for idx, val in ["*", "?", "<", ">", "/", "|", "`""] {		; Validate folder name "
		if ( inStr(newfolder, val) ) {
			badChar := true
			break
		}
	}
	
	folderChk := splitFilePath(newFolder)
	if ( !folderChk.drv || !folderChk.dir || badChar || strlen(newFolder) > 255 ) 	{		; Make sure newFolder is a valid directory string
		MsgBox("Invalid output folder")
		guiCtrl({editOutputFolder:OUTPUT_FOLDER})				; Edit reverts back to old value if new value invalid
	} else {
		OUTPUT_FOLDER := normalizePath(regExReplace(newFolder, "\\$"))	; 'Sanitize' new folder name
		LAST_OUTPUT_BROWSE_FOLDER := OUTPUT_FOLDER

		log("'" OUTPUT_FOLDER "' selected as output folder")
		SB.SetText("'" OUTPUT_FOLDER "' selected as new output folder" , 1)
		ini("write", ["OUTPUT_FOLDER", "LAST_OUTPUT_BROWSE_FOLDER"])
		refreshGUI()
		targetFocusWin := (mainAppHWND && WinExist("ahk_id " mainAppHWND)) ? "ahk_id " mainAppHWND : "ahk_id " gui1.Hwnd
		try ControlFocus("", targetFocusWin)
	}
}



; A chdman option checkbox was clicked
; ---------------------------------
checkboxOption(ctrl:="")
{
	global UI, gui1, a_GuiControl, APP_MAIN_NAME
	local ctrlName, opt, key

	guiSubmit(gui1)
	ctrlName := a_GuiControl ? a_GuiControl : (IsObject(ctrl) ? ctrl.Name : ctrl)
	opt := strReplace(ctrlName, "_checkbox")
	if ( opt == "createSubDir" && guiCtrlGet(ctrlName) == 0 && shouldWarnDisablePerJobFolder() ) {
		msg := "The selected output type '" getCurrentOutputTypeText() "' can extract many files into a single folder.`n`nContinue without creating a folder for each job?"
		if ( MsgBox(msg, APP_MAIN_NAME, 52) != "Yes" ) {
			guiCtrl({%ctrlName%:1})
			return
		}
	}
	key := guiCtrlGet(ctrlName) ? "enable":"disable"
	guiToggle(key, [opt "_dropdown", opt "_edit"])		; Enable or disable corepsnding dropdown or editfield according to checked status
	if ( UI.chdmanOpt.%opt%.HasOwnProp("masterOf") ) {										; Disable the 'slave' checkbox if masterOf is set as an option
		guiToggle(key, [UI.chdmanOpt.%opt%.masterOf "_checkbox", UI.chdmanOpt.%opt%.masterOf "_dropdown", UI.chdmanOpt.%opt%.masterOf "_edit"])
		key := UI.chdmanOpt.%opt%.masterOf "_checkbox"
		guiCtrl({%key%:0})
	}
}

shouldWarnDisablePerJobFolder()
{
	global job
	static multiFileExtractTypes := ["cue", "toc", "gdi"]

	if ( !RegExMatch(job.Cmd, "^extract") )
		return false

	outputType := StrLower(getCurrentOutputTypeText())
	return outputType ? inArray(outputType, multiFileExtractTypes) : false
}

getCurrentOutputTypeText()
{
	global job

	if ( job.HasOwnProp("selectedOutputExtTypes") && IsObject(job.selectedOutputExtTypes) && job.selectedOutputExtTypes.Length )
		return job.selectedOutputExtTypes[1]
	if ( job.HasOwnProp("OutputExtTypes") && IsObject(job.OutputExtTypes) && job.OutputExtTypes.Length )
		return job.OutputExtTypes[1]
	return ""
}
; ======================================================================================================================
; Job execution, progress, reports and cancellation
; ======================================================================================================================

; Convert job button - Start the jobs!
; ------------------------------------
buttonStartJobs()
{
	global
	local fnMsg, runCmd, thisJob, gPos, y, file, filefull, dir, x1, x2, y, cmd, x, qsToDo
	static CHDInfoFileNum
	guiSubmit(gui1)
	gui1.Opt("+ownDialogs")
	
	switch job.Cmd { 
		; Create media
		case "createcd", "createdvd", "createhd", "createraw", "createld", "extractcd", "extractdvd", "extracthd", "extractraw", "extractld":
			SB.SetText("Creating " stringUpper(job.Cmd) " work queue" , 1)
			log("Creating " stringUpper(job.Cmd) " work queue" )
			job.workQueue := createjob(job.Cmd, job.Options, job.selectedOutputExtTypes, job.selectedInputExtTypes, job.scannedFiles.%job.Cmd%)	; Create a queue (object) of files to process
			
		; Verify CHD
		case "verify":
			SB.SetText("Verifying CHD's" , 1)
			log("Starting Verify CHD's" )
			job.workQueue := createjob("verify", "", "", job.selectedInputExtTypes, job.scannedFiles.verify)
		
		; Show info for CHD
		case "info":
			SB.SetText("Info for CHD's" , 1)
			log("Getting info from CHD's" )
			job.workQueue := createjob("info", "", "", job.selectedInputExtTypes, job.scannedFiles.info)
			guiToggle("disable", "all")
			CHDInfoFileNum := 1, x1 := 20, x2:= 150, y := 20
		
			if ( gui3 )
				gui3.Destroy()
			
			gui3 := Gui('+LastFound +AlwaysOnTop +ToolWindow', "CHD Info")
			
			gui3.MarginX := 20
			gui3.MarginY := 20
			gui3.SetFont('s11 Q5 w750 c000000')
			gui3.Add('text', 'x20 y' y ' w700 vtextCHDInfoTitle', "")
			y += 40
			loop 12 {
				gui3.SetFont('s9 Q5 w700 c000000')
				gui3.Add('text', 'x' x1 ' y' y ' w100 vtextCHDInfo_' a_index, "")
				gui3.SetFont('s9 Q5 w400 c000000')
				gui3.Add('edit', 'x' x2 ' y+-13 w615 veditCHDInfo_' a_index ' readonly', "")
				y += 23
			}
			y += 30
			gui3.setFont('s9 Q5 w700 c000000')
			gui3.Add('text', 'x' x1 ' y' y, "Hunks")
			gui3.Add('text', 'x130 y' y, "Type")
			gui3.Add('text', 'x260  y' y, "Percent")
			gui3.SetFont('s9 Q5 w400 c000000')
			loop 4 {
				y += 20
				gui3.Add('text', 'x' x1 ' y' y ' w150 vtextCHDInfoHunks_' a_index, "")
				gui3.Add('text', 'x130 y' y ' w150 vtextCHDInfoType_' a_index, "")	; recombine words that were seperated with strSplit  
				gui3.Add('text', 'x260 y' y ' w150 vtextCHDInfoPercent_' a_index, "")
			}
			y += 40
			gui3.setFont('s9 Q5 w700 c000000')
			gui3.Add('text', 'x' x1 ' y' y ' w150', "Metadata")
			gui3.setFont('s8 Q5 w400 c000000', "Consolas")
			y += 20
			gui3.Add('edit', 'x' x1 ' y' y ' w750 h250 veditMetadata readonly', "")
			y+= 270
			gui3.setFont('s9 Q5 w400 c000000')
			ctrl := gui3.Add('button', 'x20 w120 y' y ' h30 vbuttonCHDInfoLeft disabled', "< Prev file")
			ctrl.OnEvent("Click", selectCHDInfo)
	applyCtrlHoverTip(ctrl, "Show the previous CHD in the info list")
			y+=5
			ctrl := gui3.Add('dropdownlist', 'x160 y' y ' w475 vdropdownCHDInfo altsubmit', getDDCHDInfoList())
			ctrl.OnEvent("Change", selectCHDInfo)
	applyCtrlHoverTip(ctrl, "Choose which CHD to show in the info window")
			try ctrl.Choose(1)
			y-=5
			gui3.setFont('s9 Q5 w400 c000000')
			ctrl := gui3.Add('button', 'x+15 y' y ' w120 h30 vbuttonCHDInfoRight',  " Next file >")
			ctrl.OnEvent("Click", selectCHDInfo)
	applyCtrlHoverTip(ctrl, "Show the next CHD in the info list")
			gui3.setFont('s12 Q5 w750 c000000')
			gui3.Add('text', 'x285 y350 w200 center border vtextCHDInfoLoading', "`nLoading...`n")
			gui3.setFont('s9 Q5 w400 c000000')

			gui3.OnEvent("Close", guiClose3)
			gui3.Show('AutoSize')
			selectCHDInfo(gui3["dropdownCHDInfo"])

		case "addMeta":
		case "delMeta":
		default:
	}
	
	if ( !job.workQueue || job.workQueue.Length == 0 ) {																	; Nothing to do!
		log("No jobs found in the work queue")
		MsgBox("No jobs in the work queue!", "Error", 16)
		return
	}

	job.workTally := {}
	job.availPSlots := []
	job.msgData := []
	job.parseReport := []
	job.allReport := ""
	job.halted := false
	job.started := false
	job.workTally := {started:0, total:job.workQueue.Length, success:0, cancelled:0, skipped:0, withError:0, finished:0, haltedMsg:"", totalFileStartSize:0, totalFileFinishSize:0}		; Set job variables
	job.workQueueSize := (job.workTally.total < JOB_QUEUE_SIZE)? job.workTally.total : JOB_QUEUE_SIZE											; If number of jobs is less then queue count, only display those progress bars
	toggleMainMenu("hide")																									; Hide main menu bar (selecting menu when running jobs stops messages from being receieved from threads)
	guiToggle("disable", "all")																								; Disable all controls while job is in progress
	guiToggle(["hide", "disable"], "buttonStartJobs")
	guiToggle(["show", "enable"], "buttonCancelAllJobs")
	
	gPos := guiGetCtrlPos("groupboxJob")
	anchorBottom := gPos.y + gPos.h
	gPos := guiGetCtrlPos("groupboxFileOptions")
	if ( gPos.h > 0 )
		anchorBottom := Max(anchorBottom, gPos.y + gPos.h)
	gPos := guiGetCtrlPos("groupboxOptions")
	if ( gPos.h > 0 )
		anchorBottom := Max(anchorBottom, gPos.y + gPos.h)
	y := anchorBottom + 25																									; Assign y value below options section
	guiCtrl("moveDraw", {groupBoxProgress:"x5 y" (anchorBottom + 5) " h" job.workQueueSize*25 + 60})						; Move and resize progress groupbox
	guiCtrl("moveDraw", {progressAll:"y" y, progressTextAll: "y" y+4}) 														; Set All Progress bar and it's text Y position
	guiCtrl( {progressAll:0, progressTextAll:"0 jobs of " job.workTally.total " completed - 0%"})
	guiToggle(["show", "enable"], ["groupBoxProgress", "progressAll", "progressTextAll"])									; Show total progress bar
	y += 35
	
	loop job.workQueueSize {																								; Show individua progress bars																
		key := "progress" a_index
		key2 := "progressText" a_index
		key3 := "progressCancelButton" a_index
		guiCtrl("moveDraw", {%key%:"y" y, %key2%:"y" y+4, %key3%:"y" y})	; Move the progress bars into place
		y += 25
		key := "progress" a_index
		key2 := "progressText" a_index
		guiCtrl({%key%:0, %key2%: ""})														; Clear the bars text and zero out percentage
		guiToggle(["enable", "show"],["progress" a_index, "progressText" a_index, "progressCancelButton" a_index])			; Enable and show job progress bars
		job.availPSlots.push(a_index)																						; Add available progress slots to queue
		job.msgData.Push({status:"finished", timeout:a_TickCount})															; Seed per-slot data so v2 array indexing remains valid																					
	}
	
	gui1.Show('autosize')																									; Resize main window to fit progress bars
		
	log(job.workTally.total " " stringUpper(job.Cmd) " jobs starting ...")
	SB.SetText(job.workTally.total " " stringUpper(job.Cmd) " jobs started" , 1)

	job.started := true
	job.startTime := a_TickCount
	OnMessage(0x004A, receiveData)																					; Receive messages from threads
	
	; Main job loop
	; -----------------------------------------------------
	loop { 																												; Loop while # of finished jobs are less then job.workTally.total
		if ( job.started == false || job.workTally.finished == job.workTally.total ) 		; Quit once all are finsihed or job was cancalled  --- Dont use workQueueSize.count() == 0 because that can be zero before all jobs have finsihed
			break

		if ( job.availPSlots.Length > 0 ) 	; Add any new jobs
			addJobsToQueue()	

		checkJobTimeouts() 					; Check for timeouts

			Sleep(1000)
	} 


	OnMessage(0x004A, receiveData, 0)

	; Finished all jobs			
	job.started := false
	job.endTime := a_Tickcount
	guiToggle("hide", "buttonCancelAllJobs")
	guiToggle("show", "buttonStartJobs")
	guiToggle("disable", "all")
	
	if ( job.halted ) {													; There was a fatal error that stopped any jobs from being attempted
		log("Fatal Error: " job.workTally.haltedMsg)
		SB.SetText("Fatal Error: " job.workTally.haltedMsg , 1)
		refreshGUI()
		MsgBox(job.workTally.haltedMsg "`n", "Fatal Error", 16)
	}
	else {																				
		fnMsg := "Total number of jobs attempted: " job.workTally.total "`n"
		fnMsg .= job.workTally.success ? job.FinPreTxt " sucessfully: " job.workTally.success "`n" : ""
		fnMsg .= job.workTally.cancelled ? "Jobs cancelled by the user: " job.workTally.cancelled "`n" : ""
		fnMsg .= job.workTally.skipped ? "Jobs skipped because the output file already exists: " job.workTally.skipped "`n" : ""
		fnMsg .= job.workTally.withError ? "Jobs that finished with errors: " job.workTally.withError "`n" : ""
		fnMsg .= "Total time to finish: " millisecToTime(job.endTime - job.startTime)
		SB.SetText("Jobs finished" (job.workTally.withError ? " with some errors":"!"), 1)
		log( regExReplace(strReplace(fnMsg, "`n", ", "), ", $", "") )

		; For CHD info mode, skip finish sounds/prompts and keep the info window as the only completion UI.
		if ( job.Cmd == "info" ) {
			refreshGUI()
			return
		}
	
		if ( PLAY_SONG_FINISHED == "yes" )																; Play sounds to indicate we are done
			playSound()

		doneMsg := job.workTally.withError ? "Finished with errors.`nWould you like to see a report?" : "Finished!`nWould you like to see a report?"
		if MsgBox(doneMsg, APP_MAIN_NAME, 36) == "Yes"
		{
			if ( gui5 )
				gui5.Destroy()
				
			gui5 := Gui('+LastFound +AlwaysOnTop +ToolWindow', "REPORT")
			gui5.OnEvent("Close", guiClose5)
			gui5.MarginX := 10
			gui5.MarginY := 20
			gui5.SetFont('s11 Q5 w700 c000000')
			gui5.Add('text',, job.Desc " report")
			gui5.SetFont('s10 Q5 w700 c000000')
			gui5.Add('text',, fnMsg)
			if ( instr(job.Desc, "create") > 0 ) { ; Show space savings if creating CHD
				txt := ""
				if ( job.workTally.totalFileStartSize > 0 && job.workTally.totalFileFinishSize > 0 )
					txt := "Total size before compression: " formatBytesReport(job.workTally.totalFileStartSize) "`nTotal size after compression: " formatBytesReport(job.workTally.totalFileFinishSize) "`nSpace savings: " formatBytesReport(job.workTally.totalFileStartSize-job.workTally.totalFileFinishSize) " (" round((1 - (job.workTally.totalFileFinishSize / job.workTally.totalFileStartSize))*100) "%)"
				gui5.Add('text', 'y+10', txt)
			}
			gui5.SetFont('s9 Q5 w400 c000000')
			reportEdit := gui5.Add('edit', 'readonly y+15 w800 h500', job.allReport)
			gui5.Show('autosize center')

			try reportEdit.Focus()
			try SendMessage(0x00B1, 0, 0, reportEdit) ; EM_SETSEL: clear selection
			return
		}
		else {
			guiClose5()
			return
		}
	}

	selectCHDInfo(guiCtrlObj, info:="", *)
	{
		global gui3, job
		infoCount := (job.scannedFiles.HasOwnProp("info") && IsObject(job.scannedFiles.info)) ? job.scannedFiles.info.Length : 0
		if ( infoCount < 1 ) {
			SetTimer(showCHDInfoLoading, 0)
			guiCtrl({textCHDInfoTitle:""})
			guiToggle("hide", "textCHDInfoLoading", 3)
			return
		}

		switch guiCtrlObj.Name {
			case "buttonCHDInfoLeft":
				CHDInfoFileNum--
			case "buttonCHDInfoRight":
				CHDInfoFileNum++
			case "dropdownCHDInfo":
				selIdx := gui3["dropdownCHDInfo"].Value
				if ( selIdx >= 1 )
					CHDInfoFileNum := selIdx
		}
		if ( CHDInfoFileNum < 1 )
			CHDInfoFileNum := 1
		else if ( CHDInfoFileNum > infoCount )
			CHDInfoFileNum := infoCount
		
		SetTimer(showCHDInfoLoading, -1000) 									; Show loading message and disable all elemnts while loading - only if loading is taking longer then normal
		if ( showCHDInfo(job.scannedFiles.info[CHDInfoFileNum], CHDInfoFileNum, infoCount, 3) == false ) { 	; Func returned nothing, clear the title
			guiCtrl({textCHDInfoTitle:""})
			guiToggle("hide", "textCHDInfoLoading", 3)
			SetTimer(showCHDInfoLoading, 0)
			return
		}
		SetTimer(showCHDInfoLoading, 0)
		guiToggle("hide", "textCHDInfoLoading", 3) 								; Hide loading message
		guiToggle("enable", ["dropdownCHDInfo", "buttonCHDInfoLeft", "buttonCHDInfoRight"], 3)		; Enable info elements
		guiCtrl("choose", {dropdownCHDInfo:CHDInfoFileNum}, 3)					; Choose dropdown to match newly selected item in CHD info
		try ControlFocus("ComboBox1", "ahk_id " gui3.Hwnd)								; Keep focus on dropdown to allow arrow keys to also select next and previous files
		
		if ( infoCount <= 1 ) {
			guiToggle("disable", ["buttonCHDInfoLeft", "buttonCHDInfoRight"], 3)
			return
		}

		if ( CHDInfoFileNum == 1 )												; Then disable the appropriate button according to selection number (first or last in list)
			guiToggle("disable", "buttonCHDInfoLeft", 3)						
		else if ( CHDInfoFileNum == infoCount )
			guiToggle("disable", "buttonCHDInfoRight", 3)
	}

	guiClose3(*)
	{
		global gui3
		gui3.Destroy()
		refreshGUI()
	}
}


; User closed the finish window
; ---------------------------------------------------
guiClose5(*)
{
	global gui5
	if ( IsObject(gui5) )
		gui5.Destroy()
	gui5 := ""
	refreshGUI()
}

; Cancel a single job in progress
; -------------------------------
progressCancelButton()
{
	global job, a_GuiControl
	
	if ( !a_guiControl )
		return
	pSlot := strReplace(a_guiControl, "progressCancelButton", "")
	if ( showAutoCloseConfirm("Cancel job " job.msgData[pSlot].idx " - " stringUpper(job.msgData[pSlot].cmd) ": " job.msgData[pSlot].workingTitle "?"
		, APP_MAIN_NAME
		, cancelConfirmExpired.Bind(pSlot)) == "Yes" )
		cancelJob(pSlot)
}


checkJobTimeouts()
{
	global job, TIMEOUT_SEC
	cancelTimeoutMs := Max(TIMEOUT_SEC * 1000, 15000)

	loop job.workQueueSize {			 																; Check for job timeouts for currently running jobs
		if ( !job.msgData.Has(a_index) || !IsObject(job.msgData[a_index]) )
			continue
		if ( !job.msgData[a_index].HasOwnProp("status") || !job.msgData[a_index].HasOwnProp("timeout") )
			continue
				
		if ( job.msgData[a_index].status == "finished" ) 	; Ignore finished jobs
			continue

		; Some worker paths can emit "cancelled" without a follow-up "finished" message.
		; Auto-finalize after a short grace period so the queue cannot hang.
		if ( job.msgData[a_index].status == "cancelled" ) {
			if ( (A_TickCount - job.msgData[a_index].timeout) > 2000 )
				finalizeCancelledSlot(a_index, job.msgData[a_index].HasOwnProp("log") && job.msgData[a_index].log ? job.msgData[a_index].log : "Job cancelled by user")
			continue
		}

		slotPid := (job.msgData[a_index].HasOwnProp("pid") ? job.msgData[a_index].pid : 0)
		slotChdPid := (job.msgData[a_index].HasOwnProp("chdmanPID") ? job.msgData[a_index].chdmanPID : 0)

		; If cancellation was requested and worker is now gone, finalize the slot as cancelled.
		if ( job.msgData[a_index].HasOwnProp("cancelRequested") && job.msgData[a_index].cancelRequested ) {
			cancelTick := job.msgData[a_index].HasOwnProp("cancelTick") ? job.msgData[a_index].cancelTick : job.msgData[a_index].timeout
			cancelElapsed := A_TickCount - cancelTick

			; Keep trying to close the CHDMAN child first.
			if ( slotChdPid && ProcessExist(slotChdPid) ) {
				try ProcessClose(slotChdPid)
				catch {
				}
			}
			if ( !(slotPid && ProcessExist(slotPid)) && !(slotChdPid && ProcessExist(slotChdPid)) ) {
				finalizeCancelledSlot(a_index, "Job cancelled by user")
				continue
			}

			; If cancellation hangs too long, force-close worker and finalize so UI cannot stall forever.
			if ( cancelElapsed > cancelTimeoutMs ) {
				if ( slotPid && ProcessExist(slotPid) ) {
					try ProcessClose(slotPid)
					catch {
					}
				}
				finalizeCancelledSlot(a_index, "Job cancellation forced after timeout", "`nWarning: Job cancellation timed out and was force-finished by namDHC.`n")
				continue
			}

			job.msgData[a_index].timeout := A_TickCount
			continue
		}

		if ( (slotPid && ProcessExist(slotPid)) || (slotChdPid && ProcessExist(slotChdPid)) ) {
			job.msgData[a_index].timeout := A_TickCount
			continue
		}
		
		if ( (a_TickCount - job.msgData[a_index].timeout) > (TIMEOUT_SEC*1000) ) {						; If timer counter exceeds threshold, we will assume thread is locked up or has errored out 
			
			job.msgData[a_index].status := "error"														; Update job.msgData[] with messages and send "error" flag for that job, then parse the data
			job.msgData[a_index].log := "Error: Job timed out"
			job.msgData[a_index].report := "Error: Job timed out`n`n"
			job.msgData[a_index].progress := 100
			job.msgData[a_index].progressText := "Timed out  -  " job.msgData[a_index].workingTitle
			parseData(job.msgData[a_index])
			
			cancelJob(job.msgData[a_index].pSlot) 														; And attempt to close the process associated with it
		}
	}
}

; Cancel all jobs currently running
; ---------------------------------
cancelAllJobs()
{
	global job, SB
	
	if ( job.started == false )
		return false
	
	if ( showAutoCloseConfirm("Are you sure you want to cancel all jobs?", APP_MAIN_NAME, cancelAllConfirmExpired) != "Yes" )
		return false
	
	loop job.workQueueSize {	
		cancelJob(a_index)
		Sleep(1)
	}
	
	loop job.workQueue.Length {
		thisJob := job.workQueue.removeAt(1)
		job.allReport .= "`n`n" stringUpper(thisJob.cmd) " - " thisJob.workingTitle "`n" drawLine(77) "`n"
		job.allReport .= "Job cancelled by user`n"
		job.workTally.cancelled++
		job.workTally.finished++
		
		percentAll := ceil((job.workTally.finished/job.workTally.total)*100)
		progressTextAll := job.workTally.finished " jobs of " job.workTally.total " completed "
		if ( job.workTally.withError )
			progressTextAll .= "(" job.workTally.withError " error" (job.workTally.withError>1? "s)":")")
		else
			progressTextAll .= " - " percentAll "`%"
		
		guiCtrl({progressAll:percentAll, progressTextAll:progressTextAll})
	}
	
	job.workQueue := []										; To make sure we are clear
	SB.SetText("Cancelling running jobs...", 1)
	return true	
}

cancelConfirmExpired(pSlot)
{
	global job
	if ( !job.msgData.Has(pSlot) || !IsObject(job.msgData[pSlot]) )
		return true

	slotData := job.msgData[pSlot]
	slotStatus := slotData.HasOwnProp("status") ? slotData.status : ""
	slotPid := slotData.HasOwnProp("pid") ? slotData.pid : 0
	slotChdPid := slotData.HasOwnProp("chdmanPID") ? slotData.chdmanPID : 0
	return (slotStatus == "finished") || (!(slotPid && ProcessExist(slotPid)) && !(slotChdPid && ProcessExist(slotChdPid)) && slotData.HasOwnProp("progress") && slotData.progress >= 100)
}

cancelAllConfirmExpired()
{
	global job
	if ( !job.started )
		return true

	loop job.workQueueSize {
		if ( !job.msgData.Has(a_index) || !IsObject(job.msgData[a_index]) )
			continue
		slotData := job.msgData[a_index]
		slotStatus := slotData.HasOwnProp("status") ? slotData.status : ""
		slotPid := slotData.HasOwnProp("pid") ? slotData.pid : 0
		slotChdPid := slotData.HasOwnProp("chdmanPID") ? slotData.chdmanPID : 0
		if ( slotStatus != "finished" && ((slotPid && ProcessExist(slotPid)) || (slotChdPid && ProcessExist(slotChdPid)) || (slotData.HasOwnProp("progress") && slotData.progress < 100)) )
			return false
	}
	return true
}



; User cancels job
; --------------------------------
cancelJob(pSlot)
{
	global job

	slotData := IsObject(job.msgData[pSlot]) ? job.msgData[pSlot] : 0
	if ( !slotData )
		return

	slotPid := slotData.HasOwnProp("pid") ? slotData.pid : 0
	slotStatus := slotData.HasOwnProp("status") ? slotData.status : ""
	if ( slotStatus == "finished" )
		return

	slotData.cancelRequested := true
	if ( !slotData.HasOwnProp("cancelTick") || !slotData.cancelTick )
		slotData.cancelTick := A_TickCount
	slotData.timeout := A_TickCount
	if ( slotStatus != "cancelled" )
		slotData.status := "cancelling"
	
	key := "progress" pSlot
	key2 := "progressText" pSlot
	guiCtrl({%key%:0})
	guiCtrl({%key2%:"Cancelling -  " (slotData.HasOwnProp("workingTitle") ? slotData.workingTitle : "Job " pSlot)})

	; If worker thread already ended (or never started), finalize cancel immediately.
	if ( !slotPid || !ProcessExist(slotPid) ) {
		finalizeCancelledSlot(pSlot, "Job cancelled by user")
		return
	}

	slotData.KILLPROCESS := "true"
	JSONStr := jsongo.Stringify(slotData)
	threadHwnd := WinExist("ahk_class AutoHotkey ahk_pid " slotPid)
	sent := false
	if ( threadHwnd )
		sent := sendAppMessage(JSONStr, "ahk_id " threadHwnd)
	if ( !sent && slotPid && ProcessExist(slotPid) ) {
		try ProcessClose(slotPid)
		catch {
		}
	}
}

finalizeCancelledSlot(pSlot, reason:="Job cancelled by user", reportTxt:="")
{
	global job
	if ( !job.msgData.Has(pSlot) || !IsObject(job.msgData[pSlot]) )
		return

	slotData := job.msgData[pSlot]
	if ( !slotData.HasOwnProp("pSlot") )
		slotData.pSlot := pSlot
	if ( !slotData.HasOwnProp("idx") )
		slotData.idx := pSlot

	slotData.status := "cancelled"
	slotData.log := reason
	slotData.report := reportTxt
	slotData.progress := 100
	slotData.progressText := "Cancelled -  " (slotData.HasOwnProp("workingTitle") ? slotData.workingTitle : ("Job " pSlot))
	parseData(slotData)

	slotData.status := "finished"
	slotData.report := ""
	parseData(slotData)
}

; ======================================================================================================================
; CHD info and worker-thread messaging
; ======================================================================================================================

; Create CHD info dropdown list
; -----------------------------
getDDCHDInfoList()
{
	global job
	ddCHDInfoList := []
	if ( !job.scannedFiles.HasOwnProp("info") || !IsObject(job.scannedFiles.info) )
		return ddCHDInfoList
	for _, filefull in job.scannedFiles.info
		ddCHDInfoList.Push(splitFilePath(filefull).file)
	return ddCHDInfoList
}


; Show CHD loading message
; --------------------
showCHDInfoLoading()
{
	guiToggle("disable", "all", 3) 
	guiToggle("show", "textCHDInfoLoading", 3) 								
}



; Show CHD info info seperate window
; -- grab new data JIT
; ----------------------------------
showCHDInfo(fullFileName, currNum, totalNum, guiNum:=3)
{
	global
	local a, line, file, infoLineNum := 0, compressLineNum := 0, metadataTxt := ""

	if ( !fileExist(fullFileName) )
		return false
		
	file := splitFilePath(fullFilename)
	guiToggle("enable", "textCHDInfoTitle", guiNum)	
	guiCtrl({textCHDInfoTitle:"[" (currNum && totalNum ? currNum "/" totalNum "]  " : "") file.file}, guiNum)																	; Change Title to filename
	runMsg := runCMD("`"" CHDMAN_FILE_LOC "`" info -v -i `"" fullFilename "`"", file.dir) ;"
	rawInfo := runMsg.HasOwnProp("msg") ? runMsg.msg : ""
	if ( !Trim(rawInfo) )
		rawInfo := runCHDInfoDirect(fullFilename, file.dir)
	loop parse, rawInfo, "`n"				; Loop through chdman 'info' stdOut
	{
		line := regExReplace(a_loopField, "`n|`r")																; Remove all CR/LF
		if ( !Trim(line) )
			continue
		if ( InStr(line, "Compressed Hunks of Data") )															; Skip banner/version line
			continue
		if ( inStr(line, "Metadata") ) {																		; If we find string 'Metadata' in line, we know to add text to metadata string
			line := strReplace(line, "Metadata:")																; Remove 'Metadata:' as its redundant
			metadataTxt .= trim(line, " ") "`n"
		}
		if ( inStr(line, "TRACK:") ) {																			; Finding 'TRACK:' in text informs us we are in metadata section of output
			line := trim(line, " "), line := strReplace(line, " ", " | "), line := strReplace(line, ":", ": ")  ; Fix formatting ...
			metadataTxt .= strReplace(line, ".") "`n`n"															; ... and add it to the metadata string
		}
		else if ( RegExMatch(line, "^\s*([^:]+):\s*(.+)$", &mKV) ) {												; Otherwise all data is part of file information
			infoLineNum++																						; Increase line number counter
			if ( infoLineNum <= 12 ) {
				guiToggle("enable", ["textCHDInfo_" infoLineNum, "editCHDInfo_" infoLineNum], guiNum)
				key := "textCHDInfo_" infoLineNum
				guiCtrl({%key%:trim(mKV[1], " ") ": "}, guiNum)							; Part 1 as subtitle
				key := "editCHDInfo_" infoLineNum
				guiCtrl({%key%:trim(mKV[2], " ")}, guiNum)								; Part 2 is the information itself
			}
		}
		else if ( RegExMatch(line, "^-{6,}\s+-{3,}\s+-{10,}$") )							; Compression section divider
			compressLineNum := 1																				; ... So flag it, use flag as the line counter and move to next loop
		else if ( compressLineNum ) {																					
			line := trim(line, a_space)
			if ( compressLineNum <= 4 && RegExMatch(line, "^\s*([\d,]+)\s+([\d.]+%)\s+(.+?)\s*$", &mComp) ) {
				guiToggle("enable", ["textCHDInfoHunks_" compressLineNum, "textCHDInfoType_" compressLineNum, "textCHDInfoPercent_" compressLineNum], guiNum)
				key := "textCHDInfoHunks_" compressLineNum
				guiCtrl({%key%:trim(mComp[1], a_space)}, guiNum)							; Part 1 is Hunks
				key := "textCHDInfoType_" compressLineNum
				guiCtrl({%key%:trim(mComp[3], a_space)}, guiNum)							; Part 2 is Compression Type
				key := "textCHDInfoPercent_" compressLineNum
				guiCtrl({%key%:trim(mComp[2], a_space)}, guiNum)							; Part 3 is percentage of compression
			}
			compressLineNum++																							; Add to meta line number
		}
	}

	if ( infoLineNum == 0 && Trim(rawInfo) ) {
		guiToggle("enable", "editMetadata", guiNum)
		guiCtrl({editMetadata:rawInfo}, guiNum)
	}
	else {
	guiToggle("enable", "editMetadata", guiNum)		; Enable all elements
	guiCtrl({editMetadata: metadataTxt}, guiNum)
	}
	try ControlFocus("", "ahk_id " gui3.Hwnd)
	return true
}

runCHDInfoDirect(fullFileName, workingDir:="")
{
	global CHDMAN_FILE_LOC
	tmpOut := A_Temp "\namdhc_info_" A_TickCount "_" Random(1000, 9999) ".txt"
	q := Chr(34)
	cmd := A_ComSpec . " /C " . q . q . CHDMAN_FILE_LOC . q . " info -v -i " . q . fullFileName . q . " > " . q . tmpOut . q . " 2>&1" . q

	try RunWait(cmd, workingDir ? workingDir : "", "Hide", &exitCode)
	catch
		return ""

	infoTxt := ""
	if ( FileExist(tmpOut) ) {
		try infoTxt := FileRead(tmpOut, "CP0")
		catch {
			try infoTxt := FileRead(tmpOut)
		}
		try FileDelete(tmpOut)
	}
	return infoTxt
}



; Receieve message data from thread script
; ----------------------------------------
receiveData(data1, data2, *) 
{
	JSONStr := StrGet(NumGet(data2 + 2*A_PtrSize, 0, "UPtr"),, "utf-8")
	data := main_jsonToLegacyObj(jsongo.Parse(JSONStr))
	parseData(data)
	return 1
}

; Parse recieved data from threads here
; Seperating from receiveData allows us to 'send' data locally as well
; ------------------------------------
parseData(recvData) 
{		
	global gui1, job, REMOVE_FILE_ENTRY_AFTER_FINISH, SB, LV
	if ( !IsObject(recvData) || !recvData.HasOwnProp("pSlot") || !recvData.HasOwnProp("idx") )
		return

	pSlot := recvData.pSlot
	if ( !IsObject(job.msgData[pSlot]) )
		job.msgData[pSlot] := {}
	mergeObj(recvData, job.msgData[pSlot])				; Thread messages may be partial. Merge into existing slot state.
	job.msgData[pSlot].timeout := a_Tickcount			; Reset timeout timer

	curr := job.msgData[pSlot]
	recvIdx := curr.HasOwnProp("idx") ? curr.idx : recvData.idx
	if ( recvIdx <= 0 )
		return

	recvStatus := recvData.HasOwnProp("status") ? recvData.status : ""
	recvLog := recvData.HasOwnProp("log") ? recvData.log : ""
	recvReport := recvData.HasOwnProp("report") ? recvData.report : ""
	recvProgress := recvData.HasOwnProp("progress") ? recvData.progress : ""
	recvProgressText := recvData.HasOwnProp("progressText") ? recvData.progressText : ""

	recvCmd := curr.HasOwnProp("cmd") ? curr.cmd : ""
	recvWorkingTitle := curr.HasOwnProp("workingTitle") ? curr.workingTitle : ""
	recvFromFileFull := curr.HasOwnProp("fromFileFull") ? curr.fromFileFull : ""
	currStatus := curr.HasOwnProp("status") ? curr.status : recvStatus
	cancelRequested := curr.HasOwnProp("cancelRequested") && curr.cancelRequested

	; If user requested cancel for this slot, don't let late success messages count as successful.
	if ( cancelRequested && recvStatus == "success" )
		recvStatus := "cancelled"

	if ( recvLog )
		log("Job " recvIdx " - " recvLog)
	
	if ( recvReport ) { 					; Report data was in with receieved data
		while ( job.parseReport.Length < recvIdx )
			job.parseReport.Push("")
		if ( !job.parseReport[recvIdx] ) 	; If we dont have report data assigned yet for this job number, assign it now
			job.parseReport[recvIdx] := "`n`n" stringUpper(recvCmd ? recvCmd : "job") " - " recvWorkingTitle "`n" drawLine(77) "`n"
		job.parseReport[recvIdx] .= recvReport
	}

	switch recvStatus {
		
		case "fileExists":
			if !(curr.HasOwnProp("countedFileExists") && curr.countedFileExists) {
				curr.countedFileExists := true
				job.workTally.skipped++
				SB.SetText("Job " recvIdx " skipped", 1)
			}
		
		case "error":
			if !(curr.HasOwnProp("countedError") && curr.countedError) {
				curr.countedError := true
				job.workTally.withError++
				SB.SetText("Job " recvIdx " failed", 1)
			}
			
		case "halted":
			job.halted := true
			job.started := false
			job.workTally.cancelled += job.workQueue.Length + 1		; Tally up totals
			job.workQueue := []											; Empty the work queue
			job.workTally.haltedMsg := recvLog							; Set flag and error log
			log("Fatal Error. Halted all jobs")
			
		case "success":
			if !(curr.HasOwnProp("countedSuccess") && curr.countedSuccess) && !(curr.HasOwnProp("cancelRequested") && curr.cancelRequested) {
				curr.countedSuccess := true
				job.workTally.success++
				SB.SetText("Job " recvIdx " finished successfully!", 1)
				if ( REMOVE_FILE_ENTRY_AFTER_FINISH == "yes" && recvFromFileFull && recvCmd && removeFinishedInputEntry(recvCmd, recvFromFileFull) )
					curr.removedFromList := true
			}
		
		case "cancelled":
			if !(curr.HasOwnProp("countedCancelled") && curr.countedCancelled) {
				curr.countedCancelled := true
				job.workTally.cancelled++
				SB.SetText("Job " recvIdx " cancelled", 1)
			}
		
		case "finished": 															; All jobs come here: Cancelled, Success, and errors
			if !(curr.HasOwnProp("countedFinished") && curr.countedFinished) {
				curr.countedFinished := true

				; If a transient "success" status update was missed, still clear completed input entries.
				if ( REMOVE_FILE_ENTRY_AFTER_FINISH == "yes"
					&& recvFromFileFull
					&& recvCmd
					&& !(curr.HasOwnProp("removedFromList") && curr.removedFromList)
					&& !(curr.HasOwnProp("countedFileExists") && curr.countedFileExists)
					&& !(curr.HasOwnProp("countedError") && curr.countedError)
					&& !(curr.HasOwnProp("countedCancelled") && curr.countedCancelled)
					&& !(curr.HasOwnProp("cancelRequested") && curr.cancelRequested) ) {
					if ( removeFinishedInputEntry(recvCmd, recvFromFileFull) )
						curr.removedFromList := true
				}

				startSize := recvData.HasOwnProp("fileStartSize") ? recvData.fileStartSize : (curr.HasOwnProp("fileStartSize") ? curr.fileStartSize : 0)
				finishSize := recvData.HasOwnProp("fileFinishSize") ? recvData.fileFinishSize : (curr.HasOwnProp("fileFinishSize") ? curr.fileFinishSize : 0)
				job.workTally.totalFileStartSize += startSize
				job.workTally.totalFileFinishSize += finishSize
				
				while ( job.parseReport.Length < recvIdx )
					job.parseReport.Push("")
				job.allReport .= job.parseReport[recvIdx]
				job.workTally.finished++
				percentAll := ceil((job.workTally.finished/job.workTally.total)*100)
				if ( percentAll > 100 )
					percentAll := 100
				
				progText := job.workTally.finished " jobs of " job.workTally.total " completed "
				if ( job.workTally.withError )
					progText .= "(" job.workTally.withError " error" (job.workTally.withError>1 ? "s)" : ")")
				else
					progText .= " - " percentAll "%"
					
				guiCtrl({progressAll:percentAll, progressTextAll:progText})
				
				job.availPSlots.push(pSlot) ; Add an available position for jobs
			}
	}


	; Update status bars
	busyBaseText := recvProgressText ? recvProgressText : (curr.HasOwnProp("progressText") ? curr.progressText : "")
	if ( currStatus == "unzipping" )
		progressSetBusy(pSlot, true, busyBaseText)
	else
		progressSetBusy(pSlot, false)

	if ( currStatus == "unzipping" ) {
		gui1["progress" pSlot].Value := 0
	}
	else {
		if ( recvData.HasOwnProp("progress") && recvProgress != "" ) ;guiControl 1:, "progress" pSlot, recvProgress
			gui1["progress" pSlot].Value := recvProgress
		if ( recvData.HasOwnProp("progressText") && recvProgressText )
			gui1["progressText" pSlot].Text := recvProgressText
	}

}

; ======================================================================================================================
; Queue construction and job objects
; ======================================================================================================================

; Add jobs to work queue
; ----------------------
addJobsToQueue()
{
	global job, SHOW_JOB_CONSOLE

	if ( job.workQueue.Length == 0 || job.availPSlots.Length == 0 )
		return false
	
	thisJob := job.workQueue.removeAt(1)														; Grab the first job from the work queue and assign parameters to variable
	thisJob.pSlot := job.availPSlots.removeAt(1)												; Assign the progress bar a y position from available queue

	job.msgData[thisJob.pSlot] := {pid:0, timeout:a_TickCount, status:"starting", pSlot:thisJob.pSlot, idx:thisJob.idx, cmd:thisJob.cmd, workingTitle:thisJob.workingTitle, progress:0, fileStartSize:0, fileFinishSize:0, report:""} ; Seed slot state before handshake
	
	threadArg := "threadMode" (SHOW_JOB_CONSOLE == "yes" ? " console" : "")			; "threadmode" flag tells script to run this script as a thread
	if ( A_IsCompiled )
		runCmd := "`"" a_ScriptFullPath "`" " threadArg
	else
		runCmd := "`"" A_AhkPath "`" `"" a_ScriptFullPath "`" " threadArg
	Run(runCmd, , "Hide", &pid)															; Launch hidden so the worker doesn't steal focus
	thisJob.pid := pid
	log("Starting new thread PID: " thisJob.pid " in slot: " thisJob.pSlot)

	handshakeStart := A_TickCount
	while ( thisJob.pid != job.msgData[thisJob.pSlot].pid ) {									; Wait for confirmation that msg was receieved																			
		if ( !ProcessExist(pid) ) {
			job.msgData[thisJob.pSlot].status := "error"
			job.msgData[thisJob.pSlot].log := "Error: Thread process exited during handshake"
			job.msgData[thisJob.pSlot].report := "Error: Thread process exited during handshake`n`n"
			job.msgData[thisJob.pSlot].progress := 100
			job.msgData[thisJob.pSlot].progressText := "Thread startup failed  -  " thisJob.workingTitle
			parseData(job.msgData[thisJob.pSlot])
			job.msgData[thisJob.pSlot].status := "finished"
			job.msgData[thisJob.pSlot].progress := 100
			job.msgData[thisJob.pSlot].report := ""
			parseData(job.msgData[thisJob.pSlot])
			return false
		}
		if ( (A_TickCount - handshakeStart) > 10000 ) {
			job.msgData[thisJob.pSlot].status := "error"
			job.msgData[thisJob.pSlot].log := "Error: Timed out while handshaking with worker thread"
			job.msgData[thisJob.pSlot].report := "Error: Timed out while handshaking with worker thread`n`n"
			job.msgData[thisJob.pSlot].progress := 100
			job.msgData[thisJob.pSlot].progressText := "Thread handshake timed out  -  " thisJob.workingTitle
			parseData(job.msgData[thisJob.pSlot])
			job.msgData[thisJob.pSlot].status := "finished"
			job.msgData[thisJob.pSlot].progress := 100
			job.msgData[thisJob.pSlot].report := ""
			parseData(job.msgData[thisJob.pSlot])
			try ProcessClose(pid)
			return false
		}
		msg := jsongo.Stringify(thisJob)
		threadHwnd := WinExist("ahk_class AutoHotkey ahk_pid " pid)
		if ( threadHwnd )
			sendAppMessage(msg, "ahk_id " threadHwnd)
		Sleep(250)	
	}
	
	job.msgData[thisJob.pSlot].timeout := a_TickCount											; Set inital timeout time
}

main_jsonToLegacyObj(val)
{
	if ( val is Map ) {
		obj := {}
		for key, item in val
			obj.%key% := main_jsonToLegacyObj(item)
		return obj
	}
	if ( val is Array ) {
		arr := []
		for _, item in val
			arr.Push(main_jsonToLegacyObj(item))
		return arr
	}
	return val
}





; Create  or add to the input files queue (return a work queue)
; -------------------------------------------------------
createJob(command, theseJobOpts, outputExts:="", inputExts:="", inputFiles:="") 
{
	global
	local idx, idx2, obj, thisOpt, optVal, splitFromFile, splitOptFile, toExt, q, fromFileFull, outputFolder, outputParentVal, resolvedOutputParent
	local wCount := 0, wQueue := [], dupFound := {}, cmdOpts := "", PID := dllCall("GetCurrentProcessId"), renameDup := 0
	
	guiSubmit(gui1)
	outputParentVal := guiCtrlGet("outputparent_checkbox") ? Trim(guiCtrlGet("outputparent_edit")) : ""
	
	for idx, thisOpt in (isObject(theseJobOpts) ? theseJobOpts : [])								; Parse through supplied Options associated with job
	{
		if ( guiCtrlGet(thisOpt.name "_checkbox", 1) == 0 )											; Skip if the checkbox is not checked
			continue
		optVal := ""
		if ( thisOpt.HasOwnProp("editField") )
			optVal := guiCtrlGet(thisOpt.name "_edit")
		else if ( thisOpt.HasOwnProp("dropdownOptions") ) {
			ddlCtrl := getGuiCtrl(thisOpt.name "_dropdown")
			optVal := guiCtrlGet(thisOpt.name "_dropdown")											; Get the dropdown value for the current UI.chdmanOpt
			if ( thisOpt.HasOwnProp("dropdownValues") && IsObject(thisOpt.dropdownValues) ) {
				ddlIdx := (ddlCtrl && ddlCtrl.HasOwnProp("Value")) ? ddlCtrl.Value : 0
				optVal := (ddlIdx >= 1 && ddlIdx <= thisOpt.dropdownValues.Length) ? thisOpt.dropdownValues[ddlIdx] : ""
			}
		}
		if ( thisOpt.HasOwnProp("paramString") && thisOpt.paramString ) {
			if ( thisOpt.HasOwnProp("editField") || thisOpt.HasOwnProp("dropdownOptions") ) {
				if ( optVal == "" )
					continue
				optVal := (thisOpt.HasOwnProp("useQuotes") && thisOpt.useQuotes) ? " `"" optVal "`"" : " " optVal
				cmdOpts .= " -" thisOpt.paramString . optVal
			} else
				cmdOpts .= " -" thisOpt.paramString
		}
	}

	if ( command == "createraw" ) {
		if ( !inStr(cmdOpts, " -hs ") ) {
			optVal := guiCtrlGet("hunksize_edit")
			optVal := optVal ? optVal : UI.chdmanOpt.hunkSize.editField
			cmdOpts .= optVal ? " -hs " optVal : ""
		}
		if ( !inStr(cmdOpts, " -us ") && !inStr(cmdOpts, " -op ") ) {
			optVal := guiCtrlGet("unitsize_edit")
			optVal := optVal ? optVal : UI.chdmanOpt.unitSize.editField
			cmdOpts .= optVal ? " -us " optVal : ""
		}
	}
	
	for idx, fromFileFull in (isObject(inputFiles) ? inputFiles : [inputFiles]) {
		splitFromFile := splitFilePath(fromFileFull)
		
		for idx, toExt in (isObject(outputExts) ? outputExts : [outputExts]) {
			
			outputFolder := USE_INPUTFOLDER_AS_OUTPUT ? splitFromFile.dir : OUTPUT_FOLDER
			
			q := {}				
			q.idx				:= wQueue.Length + 1
			q.id				:= command q.idx
			q.hostPID			:= PID
			q.cmd				:= command
			q.cmdOpts			:= cmdOpts
			q.inputFileTypes    := isObject(inputExts) ? inputExts : [inputExts]
			q.deleteInputDir	:= guiCtrlGet("deleteInputDir_checkbox")
			q.deleteInputFiles 	:= guiCtrlGet("deleteInputFiles_checkbox")
			q.keepIncomplete 	:= guiCtrlGet("keepIncomplete_checkbox")
			q.outputFolderIsPerJob := guiCtrlGet("createSubDir_checkbox") ? true : false
			q.workingDir		:= splitFromFile.dir
			q.fromFileExt		:= splitFromFile.ext
			q.fromFile			:= splitFromFile.file
			q.fromFileNoExt 	:= splitFromFile.noExt
			q.fromFileFull		:= fromFileFull
			if ( command != "verify" && command != "info" ) {
				q.toFileNoExt	:= outputExts.Length>1 ? splitFromFile.noExt " (" stringUpper(toExt) ")" : splitFromFile.noExt									; For the target file, we use the same base filename as the source
				q.outputFolder	:= q.outputFolderIsPerJob ? outputFolder "\" q.toFileNoExt : outputFolder
				q.toFileExt 	:= toExt
				q.toFile		:= q.toFileNoExt "." toExt
				q.toFileFull	:= q.outputFolder "\" q.toFileNoExt "." toExt
				; If a duplicate filename was found (ie - 'D:\folder\gameX.chd' and 'C:\folderA\gameX.chd' would both output 'gameX.cue, gameX.bin') ...
				; .. we will suffix a number to the filename gameX-1.chd and gameX-2.chd
				for idx, obj in wQueue {
					if ( obj.toFileFull == q.toFileFull ) {
						dupFound[q.toFileFull] ? dupFound[q.toFileFull]++ : dupFound[q.toFileFull] := 2
						if ( renameDup < 2 ) {
							SetTimer(changeMsgBoxButtons, 50)
							msg := MsgBox("A duplicate conversion was found.`n`n The file `"" q.fromFileFull "`"  would create a " stringUpper(toExt) " file that has the same name as another file already in the job queue.`n`nSelect [YES] to rename the output file `"" q.toFile "`" to `"" obj.toFile " [#" . dupFound[q.toFileFull] . "]`"`n`nSelect [NO] to skip this job", "Duplicate filename", 35)
							if ( msg == "Cancel" ) 	; Rename All 
								renameDup := 2
							if ( msg == "Yes" )		; Rename
								renameDup := 1
							if ( msg == "No" ) 		; Skip
								renameDup := 0
						}
						if ( renameDup ) {
							q.toFileNoExt	.= " [#" dupFound[q.toFileFull] "]"
							q.outputFolder	:= q.outputFolderIsPerJob ? outputFolder "\" q.toFileNoExt : outputFolder
							q.toFile		:= q.toFileNoExt "." toExt
							q.toFileFull	:= q.outputFolder "\" q.toFileNoExt "." toExt
						}
						else {
							dupFound[q.toFileFull]--
							q := {}
						}
						break
					}
				}
				if ( ObjOwnPropCount(q) > 0 && outputParentVal ) {
					splitOptFile := splitFilePath(outputParentVal)
					if ( !splitOptFile.dir && !splitOptFile.drv ) {
						resolvedOutputParent := q.outputFolder "\" outputParentVal
						q.cmdOpts := StrReplace(q.cmdOpts, " -op `"" outputParentVal "`"", " -op `"" resolvedOutputParent "`"")
					}
				}
			}
			if ( ObjOwnPropCount(q) > 0 ) {
				q.workingTitle 	:= (q.HasOwnProp("toFile") && q.toFile) ? q.toFile : q.fromFile
				wQueue.push(q) ; Push data to array
			}
		}
	}

	changeMsgBoxButtons(*)
	{
		if ( !winExist("Duplicate filename") )
			return 
		SetTimer(changeMsgBoxButtons, 0) 
		WinActivate("Duplicate filename")
		ControlSetText("&Rename", "Button1", "Duplicate filename")
		ControlSetText("&Skip", "Button2", "Duplicate filename")
		ControlSetText("Rename &All", "Button3", "Duplicate filename")
	}
	return wQueue
}

; ======================================================================================================================
; Main GUI creation, menus, logging and app settings
; ======================================================================================================================

; Create the main GUI
; -------------------
createMainGUI() 
{
	global
	local idx, key, opt, optName, obj, btn, ctrl, array := [], ddList := ""
	
	gui1 := Gui('+OwnDialogs', APP_MAIN_NAME)
	gui1.OnEvent("Close", (*) => quitApp())
	gui1.OnEvent("DropFiles", legacyGuiDropFiles.Bind(guiDropFiles))
	
	if ( APP_NO_DPI_SCALE == "yes" )
		gui1.Opt('-DPIScale') 		; hacky workaround to at least get the options crammed in view - for those folks who use higher DPI settings
	
	gui1.Add('button', 'hidden default h0 w0 y0 y0', "").OnEvent("Click", legacyGuiEvent.Bind(editOutputFolder))		; For output edit field (default button
	
	SB := gui1.Add('statusBar',, "namDHC")
	SB.SetParts(640, 175)
	SB.SetText('  namDHC v' CURRENT_VERSION ' for CHDMAN', 2)

	gui1.Add('groupBox', 'x5 w800 h450 vgroupboxJob', "Jobs")

	gui1.Add('text', 'x15 y30', "Job type:")
	
	for key, obj in UI.dropdowns.job.OwnProps()					; Get job dropdown list from object-dictionary
		array.Push(obj.desc)		
	ctrl := gui1.Add('dropDownList', 'x+5 y28 w200 Choose1 vdropdownJob', array)
	ctrl.OnEvent("Change", legacyGuiEvent.Bind(selectJob))
	applyCtrlHoverTip(ctrl, "Choose whether to create, extract, verify, or read CHD info")

	gui1.Add('text', 			'x+30 y30', "Media type:")
	ctrl := gui1.Add('dropDownList',	'x+5 y28 w200 vdropdownMedia') 	; Media dropdown will be populated when selcting a job
	ctrl.OnEvent("Change", legacyGuiEvent.Bind(selectMedia))
	applyCtrlHoverTip(ctrl, "Choose the media type for the current job")

	gui1.Add('text', 			'x15 y65', "Input files")
	
	ctrl := gui1.Add('button', 		'x15 y83 w80 h22 vbuttonAddFiles', "Add files")
	ctrl.OnEvent("Click", (ctrlObj, *) => addFolderFiles(false, ctrlObj.Name))
	applyCtrlHoverTip(ctrl, "Add one or more files to the job list")
	ctrl := gui1.Add('button', 		'x+5 y83 w90 h22 vbuttonAddFolder', "Add a folder")
	ctrl.OnEvent("Click", (ctrlObj, *) => addFolderFiles(false, ctrlObj.Name))
	applyCtrlHoverTip(ctrl, "Scan a folder and add matching files to the job list")
	
	ctrl := gui1.Add('button', 			'x560 y83 w233 h22 vbuttonInputExtType', "Select input file types")
	ctrl.OnEvent("Click", legacyGuiEvent.Bind(buttonExtSelect))
	applyCtrlHoverTip(ctrl, "Choose which input file types namDHC should include for this job")
	
	LV := gui1.Add('listView', 'x15 y110 w778 h153 vlistViewInputFiles altsubmit', ["File"])
	LV.OnEvent("ItemSelect", legacyGuiEvent.Bind(listViewInputFiles))
	
	ctrl := gui1.Add('button', 		'x15 y267 w90 vbuttonSelectAllInputFiles', "Select all")
	ctrl.OnEvent("Click", legacyGuiEvent.Bind(selectInputFiles))
	applyCtrlHoverTip(ctrl, "Select every file in the current list")
	ctrl := gui1.Add('button', 		'x+5 y267 w90 vbuttonClearInputFiles', "Unselect")
	ctrl.OnEvent("Click", legacyGuiEvent.Bind(selectInputFiles))
	applyCtrlHoverTip(ctrl, "Clear the current file selection")
	ctrl := gui1.Add('button', 		'x+20 y267 w120 vbuttonRemoveInputFiles', "Remove selection")
	ctrl.OnEvent("Click", legacyGuiEvent.Bind(selectInputFiles))
	applyCtrlHoverTip(ctrl, "Remove the selected entries from the file list")

	gui1.Add('text', 			'x15 y305', "Output Folder")
	ctrl := gui1.Add('button', 		'x15 y324 w150 h24 vbuttonBrowseOutput', "Select output folder")
	ctrl.OnEvent("Click", (ctrlObj, *) => checkNewOutputFolder(ctrlObj.Name))
	applyCtrlHoverTip(ctrl, "Choose where finished output files will be saved")


	ctrl := gui1.Add('button',		'x560 y324 w233 h24 vbuttonOutputExtType', "Select output file type")
	ctrl.OnEvent("Click", legacyGuiEvent.Bind(buttonExtSelect))
	applyCtrlHoverTip(ctrl, "Choose which output file type this job should create")
	ctrl := gui1.Add('edit', 		'x15 y360 w778 veditOutputFolder', OUTPUT_FOLDER)
	ctrl.OnEvent("LoseFocus", legacyGuiEvent.Bind(editOutputFolder))
	ctrl := gui1.Add('checkbox',	'x15 y390 w200 -wrap checked' USE_INPUTFOLDER_AS_OUTPUT ' vUSE_INPUTFOLDER_AS_OUTPUT', "Save output in input file folders")
	ctrl.OnEvent("Click", legacyGuiEvent.Bind(checkUseInputFolderAsOutput))
	applyCtrlHoverTip(ctrl, "Save output beside each source file instead of using the output folder above")

	ctrl := gui1.Add('button',		'x320 y410 w160 h35 vbuttonStartJobs', "Start all jobs!")
	ctrl.OnEvent("Click", legacyGuiEvent.Bind(buttonStartJobs))
	applyCtrlHoverTip(ctrl, "Build the queue and start all listed jobs")
	
	ctrl := gui1.Add('button',		'hidden x320 y395 w160 h35 vbuttonCancelAllJobs', "CANCEL ALL JOBS")
	ctrl.OnEvent("Click", legacyGuiEvent.Bind(cancelAllJobs))
	applyCtrlHoverTip(ctrl, "Stop all running and queued jobs")

	gui1.SetFont('Q5 s9 w400 c000000')
	gui1.Add('groupBox', 	'x5 w800 y460 h0 hidden vgroupboxFileOptions', "File Options")
	gui1.Add('groupBox', 	'x5 w800 y460 h0 hidden vgroupboxOptions', "CHDMAN Options")		; Position and height will be set in refreshGUI()
	gui1.SetFont('Q5 s9 w400 c000000')

	for key, opt in UI.chdmanOpt.OwnProps()
	{
		if ( opt.HasOwnProp("hidden") && opt.hidden == true )
			continue
		optName := opt.name
		ctrl := gui1.Add('checkbox',		'hidden w200 -wrap v' optName '_checkbox')	; Options are moved to their positions when refreshGUI(true) is called
		ctrl.OnEvent("Click", legacyGuiEvent.Bind(checkboxOption))
		applyCtrlHoverTip(ctrl, opt)
		editCtrl := gui1.Add('edit',			'hidden w165 v' optName '_edit')
		applyCtrlHoverTip(editCtrl, opt)
		ddlCtrl := gui1.Add('dropdownList',  'hidden w165 altsubmit v' optName '_dropdown')				; ... so we can use for dropdown list to place at same location (default is hidden)
		applyCtrlHoverTip(ddlCtrl, opt)
	}

}

checkUseInputFolderAsOutput()
{
	global USE_INPUTFOLDER_AS_OUTPUT, gui1, a_GuiControl
	guiSubmit(gui1)
	USE_INPUTFOLDER_AS_OUTPUT := gui1["USE_INPUTFOLDER_AS_OUTPUT"].Value

	if ( a_GuiControl == "USE_INPUTFOLDER_AS_OUTPUT" ) ; If checked/unchecked by user
		ini("write", "USE_INPUTFOLDER_AS_OUTPUT")

	if ( USE_INPUTFOLDER_AS_OUTPUT )
		guiToggle("disable", ["buttonBrowseOutput", "editOutputFolder"])
	else
		guiToggle("enable", ["buttonBrowseOutput", "editOutputFolder"])

}

; Create GUI progress bar section
; -------------------------------
createProgressBars()
{
	global
	local btn
	
	gui1.Add('groupBox', 'w800 vgroupBoxProgress', "Progress")

	gui1.setFont('Q5 s9 w700 cFFFFFF')
	gui1.Add('progress', 'hidden x20 w770 h22 backgroundAAAAAA vprogressAll cgreen', 0)		; Progress bars y values will be determined with refreshGUI()
	gui1.Add('text', 'hidden x30 w750 h22 +backgroundTrans -wrap vprogressTextAll', 0)

	loop JOB_QUEUE_SIZE_LIMIT {																; Draw but hide all progress bars - we will only show what is called for later
		gui1.Add('progress', 'hidden x20 w740 h22 backgroundAAAAAA vprogress' a_index ' c17A2B8', 0)	
		gui1.Add('text', 'hidden x30 w720 h22 +backgroundTrans -wrap vprogressText' a_index, "")
		ctrl := gui1.Add('button', 'hidden x+15 w25 vprogressCancelButton' a_index, "X")
		ctrl.OnEvent("Click", legacyGuiEvent.Bind(progressCancelButton))
				applyCtrlHoverTip(ctrl, "Cancel this individual job")
	}
}
 

; Create GUI Menus
; ----------------
createMenus() 
{
	global UI, JOB_QUEUE_SIZE_LIMIT, JOB_QUEUE_SIZE
	UI.menuObj := {}
	UI.menuObj.MainMenu := MenuBar()
	UI.menuObj.SubSettingsConcurrently := Menu()

	loop JOB_QUEUE_SIZE_LIMIT
		UI.menuObj.SubSettingsConcurrently.Add(a_index, legacyMenuEvent.Bind("menuSelected", "SubSettingsConcurrently"))
	UI.menuObj.SubSettingsConcurrently.Check(JOB_QUEUE_SIZE "&")				; Select current jobQueue number

	loop UI.menu.namesOrder.Length {
		menuName := UI.menu.namesOrder[a_index]
		menuArray := UI.menu.%menuName%
		thisMenuName := menuName "Menu"
		UI.menuObj.%thisMenuName% := Menu()
		
		loop menuArray.Length {
			menuItem :=  menuArray[a_index]
			callback := menuItem.gotolabel == ":SubSettingsConcurrently"
				? UI.menuObj.SubSettingsConcurrently
				: legacyMenuEvent.Bind(menuItem.gotolabel, thisMenuName)
			UI.menuObj.%thisMenuName%.Add(menuItem.name, callback)
		
			if ( menuItem.saveVar ) {
				saveVar := menuItem.saveVar
				if ( namedSetting(saveVar) == "yes" )
					UI.menuObj.%thisMenuName%.Check(menuItem.name)
				else
					UI.menuObj.%thisMenuName%.Uncheck(menuItem.name)
			}
		}
		UI.menuObj.MainMenu.Add(menuName, UI.menuObj.%thisMenuName%)
	}
	gui1.MenuBar := UI.menuObj.MainMenu
	
	UI.menuObj.InputExtTypes := Menu()				; Input & Output extension dummy menus to populate with refreshGUI() later				
	UI.menuObj.OutputExtTypes := Menu()
	
	A_TrayMenu.Delete()						; Remove default options in tray icon
	A_TrayMenu.Add("E&xit", (*) => quitApp())
}



; Show or hide main menu
; -------------------------------
toggleMainMenu(showOrHide:="show")
{
	global mainAppHWND
	static visble := false, hMenu := 0
	
	if ( !hMenu )
		hMenu := DllCall("GetMenu", "uint", mainAppHWND)			; Save menu to retrieve later
	
	if ( showOrHide == "show" && !visble ) {
		dllCall("SetMenu", "uint", mainAppHWND, "uint", hMenu)
		visble := true
	}
	else if ( showOrHide == "hide" && visble ) {
		dllCall("SetMenu", "uint", mainAppHWND, "uint", 0)
		visble := false
	}
}


; Show or hide the verbose window
; -------------------------------
showVerbose(show:="yes")
{
	global 

	opts := 'w' APP_VERBOSE_WIN_WIDTH ' h' APP_VERBOSE_WIN_HEIGHT ' x' APP_VERBOSE_WIN_POS_X ' y' APP_VERBOSE_WIN_POS_Y
	if ( !gui2 ) {
		gui2 := Gui('-sysmenu +resize', APP_VERBOSE_NAME)
		gui2.OnEvent("Size", GuiSize2)
		gui2.OnEvent("Close", (*) => gui2.Hide())
		
		gui2.MarginX := 5
		gui2.MarginY := 10
		gui2.Add('edit', 'w' APP_VERBOSE_WIN_WIDTH-10 ' h' APP_VERBOSE_WIN_HEIGHT-20 ' readonly veditVerbose')
	}
	
	if ( show == "yes" ) {
		gui2.Show(opts) 
		gui2["editVerbose"].Focus()
		PostMessage(0x115, 7, 0, gui2["editVerbose"])
	}
	else
		gui2.Hide()
}


; Log messages and send to verbose window
; --------------------------------------
log(newMsg:="", newline:=true, clear:=false, timestamp:=true) 
{
	global
	
	if ( !newMsg ) 
		return false
	
	newMsg := timestamp ? "[" a_Hour ":" a_Min ":" a_Sec "]  " newMsg : newMsg
	local msg := clear? newMsg : guiCtrlGet("editVerbose", 2) . newMsg
	guiCtrl({editVerbose:msg (newline? "`n" : "")}, 2)
	PostMessage(0x115, 7, 0, gui2["editVerbose"])	; Scroll to bottom of log
}



namedSetting(varName, newVal:="", setVal:=false)
{
	global

	switch varName {
		case "JOB_QUEUE_SIZE":
			if ( setVal )
				JOB_QUEUE_SIZE := newVal
			return JOB_QUEUE_SIZE
		case "OUTPUT_FOLDER":
			if ( setVal )
				OUTPUT_FOLDER := newVal
			return OUTPUT_FOLDER
		case "LAST_INPUT_BROWSE_FOLDER":
			if ( setVal )
				LAST_INPUT_BROWSE_FOLDER := newVal
			return LAST_INPUT_BROWSE_FOLDER
		case "LAST_OUTPUT_BROWSE_FOLDER":
			if ( setVal )
				LAST_OUTPUT_BROWSE_FOLDER := newVal
			return LAST_OUTPUT_BROWSE_FOLDER
		case "SHOW_JOB_CONSOLE":
			if ( setVal )
				SHOW_JOB_CONSOLE := newVal
			return SHOW_JOB_CONSOLE
		case "SHOW_VERBOSE_WINDOW":
			if ( setVal )
				SHOW_VERBOSE_WINDOW := newVal
			return SHOW_VERBOSE_WINDOW
		case "PLAY_SONG_FINISHED":
			if ( setVal )
				PLAY_SONG_FINISHED := newVal
			return PLAY_SONG_FINISHED
		case "REMOVE_FILE_ENTRY_AFTER_FINISH":
			if ( setVal )
				REMOVE_FILE_ENTRY_AFTER_FINISH := newVal
			return REMOVE_FILE_ENTRY_AFTER_FINISH
		case "APP_NO_DPI_SCALE":
			if ( setVal )
				APP_NO_DPI_SCALE := newVal
			return APP_NO_DPI_SCALE
		case "USE_INPUTFOLDER_AS_OUTPUT":
			if ( setVal )
				USE_INPUTFOLDER_AS_OUTPUT := newVal
			return USE_INPUTFOLDER_AS_OUTPUT
		case "APP_MAIN_WIN_POS_X":
			if ( setVal )
				APP_MAIN_WIN_POS_X := newVal
			return APP_MAIN_WIN_POS_X
		case "APP_MAIN_WIN_POS_Y":
			if ( setVal )
				APP_MAIN_WIN_POS_Y := newVal
			return APP_MAIN_WIN_POS_Y
		case "APP_VERBOSE_WIN_WIDTH":
			if ( setVal )
				APP_VERBOSE_WIN_WIDTH := newVal
			return APP_VERBOSE_WIN_WIDTH
		case "APP_VERBOSE_WIN_HEIGHT":
			if ( setVal )
				APP_VERBOSE_WIN_HEIGHT := newVal
			return APP_VERBOSE_WIN_HEIGHT
		case "APP_VERBOSE_WIN_POS_X":
			if ( setVal )
				APP_VERBOSE_WIN_POS_X := newVal
			return APP_VERBOSE_WIN_POS_X
		case "APP_VERBOSE_WIN_POS_Y":
			if ( setVal )
				APP_VERBOSE_WIN_POS_Y := newVal
			return APP_VERBOSE_WIN_POS_Y
		case "CHECK_FOR_UPDATES_STARTUP":
			if ( setVal )
				CHECK_FOR_UPDATES_STARTUP := newVal
			return CHECK_FOR_UPDATES_STARTUP
		case "SEVENZIP_CUSTOM_PATH":
			if ( setVal )
				SEVENZIP_CUSTOM_PATH := newVal
			return SEVENZIP_CUSTOM_PATH
	}
	return ""
}


; Read or write to ini file
; -------------------------
ini(job:="read", var:="") 
{
	global
	local varsArry := isObject(var)? var : [var], idx, varName
	
	if ( varsArry[1] == "" )
		return false

	for idx, varName in varsArry {
		if ( !IsSet(varName) || varName == "" )
			continue
		if ( job == "read" ) {
			defaultVar := namedSetting(varName)
			newVal := IniRead(APP_MAIN_NAME ".ini", "Settings", varName, defaultVar)
			namedSetting(varName, (newVal == "ERROR" || newVal == "") ? defaultVar : newVal, true)
		}
		else if ( job == "write" ) {
			IniWrite(namedSetting(varName), APP_MAIN_NAME ".ini", "Settings", varName)
			;log("Saved " varName " with value " %varName%)
		}
	}
}


; Play a sound
; ------------
playSound() 
{
	; Use a short native Windows notification sound (Win10/11).
	; MB_ICONASTERISK (0x40) maps to the configured "Asterisk/Notification" system sound.
	if ( !DllCall("User32.dll\MessageBeep", "UInt", 0x40) )
		SoundBeep(900, 120)
}



; Send data across script instances
; -------------------------------------------------------
sendAppMessage(stringToSend, targetScriptTitle) 
{
  utf8 := Buffer(StrPut(stringToSend, "utf-8"), 0)
  SizeInBytes := StrPut(stringToSend, utf8, "utf-8")
  CopyDataStruct := Buffer(3*A_PtrSize, 0)
  NumPut("UPtr", SizeInBytes, CopyDataStruct, A_PtrSize)
  NumPut("UPtr", utf8.Ptr, CopyDataStruct, 2*A_PtrSize)
  Prev_DetectHiddenWindows := A_DetectHiddenWindows
  Prev_TitleMatchMode := A_TitleMatchMode
  DetectHiddenWindows(True)
  SetTitleMatchMode(2)
  try
	rtn := SendMessage(0x4A, 0, CopyDataStruct.Ptr, , targetScriptTitle)
  catch
	rtn := 0
  DetectHiddenWindows(Prev_DetectHiddenWindows)
  SetTitleMatchMode(Prev_TitleMatchMode)
  return rtn
}


; runCMD v0.94 by SKAN on D34E/D37C @ autohotkey.com/boards/viewtopic.php?t=74647
; Based on StdOutToVar.ahk by Sean @ autohotkey.com/board/topic/15455-stdouttovar      

runCMD(CmdLine, workingDir:="", codepage:="CP0", Fn:="RunCMD_Output")  
{
	global RUNCMD_STATE

	if ( Type(Fn) == "String" && Fn != "" ) {
		try Fn := Func(Fn)
		catch
			Fn := 0
	}
	else if ( !(IsObject(Fn) && HasMethod(Fn, "Call")) )
		Fn := 0

	hPipeR := 0, hPipeW := 0, hProcess := 0, hThread := 0
	pid := 0
	sOutput := ""
	exitCode := -1
	carryBuf := ""
	lineNum := 1
	buf := Buffer(16384, 0)

	try {
		if !DllCall("CreatePipe", "PtrP", hPipeR, "PtrP", hPipeW, "Ptr", 0, "UInt", 0)
			throw Error("CreatePipe failed. LastError=" A_LastError)

		; Child should inherit write handle only; parent keeps read handle local.
		DllCall("SetHandleInformation", "Ptr", hPipeW, "UInt", 1, "UInt", 1) ; HANDLE_FLAG_INHERIT on
		DllCall("SetHandleInformation", "Ptr", hPipeR, "UInt", 1, "UInt", 0) ; HANDLE_FLAG_INHERIT off

		P8 := (A_PtrSize = 8)
		SI := Buffer(P8 ? 104 : 68, 0)
		NumPut("UInt", SI.Size, SI, 0)
		NumPut("UInt", 0x100, SI, P8 ? 60 : 44)       ; STARTF_USESTDHANDLES
		NumPut("Ptr", hPipeW, SI, P8 ? 88 : 60)       ; hStdOutput
		NumPut("Ptr", hPipeW, SI, P8 ? 96 : 64)       ; hStdError
		PI := Buffer(P8 ? 24 : 16, 0)

		createFlags := 0x08000000 | DllCall("GetPriorityClass", "Ptr", -1, "UInt")
		if !DllCall("CreateProcess", "Ptr", 0, "Str", CmdLine, "Ptr", 0, "Ptr", 0
			, "Int", True, "UInt", createFlags, "Ptr", 0
			, "Ptr", workingDir ? StrPtr(workingDir) : 0
			, "Ptr", SI.Ptr, "Ptr", PI.Ptr)
			throw Error("CreateProcess failed. LastError=" A_LastError)

		hProcess := NumGet(PI, 0, "Ptr")
		hThread := NumGet(PI, A_PtrSize, "Ptr")
		pid := NumGet(PI, P8 ? 16 : 8, "UInt")
		RUNCMD_STATE.PID := pid

		if ( hPipeW ) {
			DllCall("CloseHandle", "Ptr", hPipeW)
			hPipeW := 0
		}

		loop {
			bytesAvail := 0
			if ( DllCall("PeekNamedPipe", "Ptr", hPipeR, "Ptr", 0, "UInt", 0, "Ptr", 0, "UIntP", bytesAvail, "Ptr", 0) && bytesAvail > 0 ) {
				toRead := (bytesAvail > buf.Size) ? buf.Size : bytesAvail
				bytesRead := 0
				if ( DllCall("ReadFile", "Ptr", hPipeR, "Ptr", buf.Ptr, "UInt", toRead, "UIntP", bytesRead, "Ptr", 0) && bytesRead > 0 ) {
					chunk := StrGet(buf.Ptr, bytesRead, codepage)
					sOutput .= runCMD_ProcessChunk(chunk, Fn, &lineNum, pid, &carryBuf)
					continue
				}
			}

			waitState := DllCall("WaitForSingleObject", "Ptr", hProcess, "UInt", 0, "UInt")
			; WAIT_TIMEOUT (0x102) means process still running.
			if ( waitState != 0x102 ) {
				; If PID is still alive, keep waiting even if process handle appears signaled/failed.
				if ( pid && ProcessExist(pid) ) {
					Sleep(25)
					continue
				}
				loop {
					bytesAvail := 0
					if !(DllCall("PeekNamedPipe", "Ptr", hPipeR, "Ptr", 0, "UInt", 0, "Ptr", 0, "UIntP", bytesAvail, "Ptr", 0) && bytesAvail > 0)
						break
					toRead := (bytesAvail > buf.Size) ? buf.Size : bytesAvail
					bytesRead := 0
					if !(DllCall("ReadFile", "Ptr", hPipeR, "Ptr", buf.Ptr, "UInt", toRead, "UIntP", bytesRead, "Ptr", 0) && bytesRead > 0)
						break
					chunk := StrGet(buf.Ptr, bytesRead, codepage)
					sOutput .= runCMD_ProcessChunk(chunk, Fn, &lineNum, pid, &carryBuf)
				}
				if !DllCall("GetExitCodeProcess", "Ptr", hProcess, "UIntP", procExit:=0)
					procExit := -1
				exitCode := procExit
				break
			}
			Sleep(25)
		}

	}
	catch as e {
		errMsg := "runCMD_CreateProcess error: " e.Message
		if ( sOutput )
			sOutput .= "`n" errMsg "`n"
		else
			sOutput := errMsg "`n"
		exitCode := -1
	}

	sOutput .= runCMD_ProcessChunk("", Fn, &lineNum, pid, &carryBuf, true)
	RUNCMD_STATE.PID := 0
	if ( hPipeW )
		DllCall("CloseHandle", "Ptr", hPipeW)
	if ( hPipeR )
		DllCall("CloseHandle", "Ptr", hPipeR)
	if ( hThread )
		DllCall("CloseHandle", "Ptr", hThread)
	if ( hProcess )
		DllCall("CloseHandle", "Ptr", hProcess)
	return {msg:sOutput, exitcode:exitCode}
}

runCMD_ProcessChunk(text, Fn, &lineNum, pid, &carryBuf, flushTail:=false)
{
	out := ""
	if ( text != "" ) {
		for _, ch in StrSplit(text)
		{
			if ( ch == "`r" || ch == "`n" ) {
				if ( carryBuf == "" )
					continue
				line := carryBuf "`n"
				carryBuf := ""
				if ( Fn ) {
					try out .= Fn.Call(line, lineNum++, pid)
					catch
						out .= line
				}
				else
					out .= line
				continue
			}
			carryBuf .= ch
		}
	}

	if ( flushTail && carryBuf != "" ) {
		line := carryBuf "`n"
		carryBuf := ""
		if ( Fn ) {
			try out .= Fn.Call(line, lineNum++, pid)
			catch
				out .= line
		}
		else
			out .= line
	}
	return out
}
; ======================================================================================================================
; Shared utility helpers
; ======================================================================================================================

; Filesystem and path helpers
; ---------------------------------------------

; Create a folder
; ---------------------------------------------
createFolder(newFolder) 
{
	if ( fileExist(newFolder) != "D" ) {						; Folder dosent exist
		if ( !splitFilePath(newFolder).drv ) {						; No drive letter can be assertained, so it's invalid
			newFolder := false
		} else {												; Output folder is valid but dosent exist
			try DirCreate(regExReplace(newFolder, "\\$"))
			catch
				newFolder := false
		}
	}
	return newFolder											; Returns the folder name if created or it exists, or false if no folder was created
}

; GUI and menu helpers
; -----------------------------------------------

; Check if menu item has a checkmark (is checked)
; -----------------------------------------------
isMenuChecked(menuName, itemNumber)  
{
	static MIIM_STATE := 1, MFS_CHECKED := 0x8
	hMenu := UI.menuObj.%menuName%.Handle
	MENUITEMINFO := Buffer(4*4 + A_PtrSize*8, 0)
	NumPut("UInt", MENUITEMINFO.Size, MENUITEMINFO)
	NumPut("UInt", MIIM_STATE, MENUITEMINFO, 4)
	DllCall("GetMenuItemInfo", "Ptr", hMenu, "UInt", itemNumber - 1, "UInt", true, "Ptr", MENUITEMINFO.Ptr)
	return !!(NumGet(MENUITEMINFO, 4*3, "UInt") & MFS_CHECKED)
}
; Data, object, string and array helpers
; -------------------------

; Merge two objects or arrays
; -------------------------
mergeObj(sourceObj, targetObj) 
{
	targetObj := isObject(targetObj) ? targetObj : {}
	for k, v In sourceObj.OwnProps() {
		if ( isObject(v) ) {
			if ( !targetObj.HasOwnProp(k) )
				targetObj.%k% := {}
			mergeObj(v, targetObj.%k%)
		} else
			targetObj.%k% := v
	}
}

; Create processor count dropdown list
; -------------------
procCountDDList()
{
	loop EnvGet("NUMBER_OF_PROCESSORS")
		lst .= a_index "|"					
	return "|" lst "|"						; Last "|" is to select last as default
}

; Splitpath function
; -------------------
splitFilePath(inputFile) 
{
	SplitPath(inputFile, &file, &dir, &ext, &noext, &drv)
	return {full:inputFile, file:file, dir:dir, ext:ext, noext:noext, drv:drv}
}

getPathExt(pathStr)
{
	fileName := RegExReplace(pathStr, "^.*[\\/]")
	return RegExMatch(fileName, "\.([^.\\\/]+)$", &m) ? m[1] : ""
}

; Get position of a GUI control
; --------------------------------
guiGetCtrlPos(ctrl, guiNum:=1) 
{
	ctrlObj := getGuiCtrl(ctrl, guiNum)
	if ( !ctrlObj )
		return {x:0, y:0, w:0, h:0}
	ctrlObj.GetPos(&x, &y, &w, &h)
	return {x:x, y:y, w:w, h:h}
}

guiGetCtrlRectScreen(ctrl, guiNum:=1)
{
	ctrlObj := getGuiCtrl(ctrl, guiNum)
	if ( !ctrlObj )
		return {x:0, y:0, w:0, h:0}

	rect := Buffer(16, 0)
	DllCall("GetWindowRect", "Ptr", ctrlObj.Hwnd, "Ptr", rect)
	left := NumGet(rect, 0, "Int")
	top := NumGet(rect, 4, "Int")
	right := NumGet(rect, 8, "Int")
	bottom := NumGet(rect, 12, "Int")
	return {x:left, y:top, w:right-left, h:bottom-top}
}

showMenuUnderCtrl(menuObj, ctrl, guiNum:=1)
{
	ctrlPos := guiGetCtrlRectScreen(ctrl, guiNum)
	CoordMode("Menu", "Screen")
	menuObj.Show(ctrlPos.x, ctrlPos.y + ctrlPos.h)
}

getGuiObj(guiNum:=1)
{
	global gui1, gui2, gui3, gui4, gui5, gui6
	switch guiNum {
		case 1: return gui1
		case 2: return gui2
		case 3: return gui3
		case 4: return gui4
		case 5: return gui5
		case 6: return gui6
	}
	return ""
}

progressSetBusy(pSlot, enable:=true, baseText:="")
{
	global gui1
	static busyState := Map(), timerOn := false

	if ( !gui1 )
		return

	try ctrl := gui1["progress" pSlot]
	catch
		return
	if ( !ctrl )
		return

	if ( enable ) {
		if ( !busyState.Has(pSlot) )
			busyState[pSlot] := {baseText:"", dots:0}
		state := busyState[pSlot]
		if ( baseText )
			state.baseText := RegExReplace(baseText, "\.+$")
		else if ( !state.baseText )
			state.baseText := "Unzipping"
		busyState[pSlot] := state
		ctrl.Value := 0
		if ( !timerOn ) {
			SetTimer(progressBusyTick, 220)
			timerOn := true
		}
		progressBusyTick()
		return
	}

	if ( busyState.Has(pSlot) )
		busyState.Delete(pSlot)
	if ( timerOn && busyState.Count == 0 ) {
		SetTimer(progressBusyTick, 0)
		timerOn := false
	}

	progressBusyTick()
	{
		global gui1

		if ( !gui1 )
			return

		if ( busyState.Count == 0 ) {
			if ( timerOn ) {
				SetTimer(progressBusyTick, 0)
				timerOn := false
			}
			return
		}

		for thisPSlot, state in busyState {
			try thisCtrl := gui1["progress" thisPSlot]
			catch
				continue
			try textCtrl := gui1["progressText" thisPSlot]
			catch
				continue
			if ( !thisCtrl || !textCtrl )
				continue

			state.dots := Mod(state.dots, 6) + 1
			busyState[thisPSlot] := state
			thisCtrl.Value := 0
			textCtrl.Text := state.baseText . repeatString(".", state.dots)
		}
	}
}

repeatString(str, count)
{
	rtn := ""
	Loop Max(0, count)
		rtn .= str
	return rtn
}

getGuiCtrl(ctrl, guiNum:=1)
{
	guiObj := getGuiObj(guiNum)
	if ( !guiObj )
		return ""

	try return guiObj[ctrl]
	catch {
		try {
			hwnd := ControlGetHwnd(ctrl, "ahk_id " guiObj.Hwnd)
			return hwnd ? GuiCtrlFromHwnd(hwnd) : ""
		}
		catch {
			return ""
		}
	}
}

; Register concise hover help for a GUI control
; --------------------------------------------
registerCtrlHoverTip(ctrlObj, tipText)
{
	global UI

	if ( !ctrlObj || !tipText )
		return
	if ( !UI.HasOwnProp("ctrlHoverTips") || !(UI.ctrlHoverTips is Map) )
		UI.ctrlHoverTips := Map()
	UI.ctrlHoverTips[ctrlObj.Hwnd] := tipText
}

; Apply hover help from a string or object definition
; --------------------------------------------------
applyCtrlHoverTip(ctrlObj, tipSource)
{
	if ( !ctrlObj || !tipSource )
		return

	if ( IsObject(tipSource) ) {
		if ( tipSource.HasOwnProp("helpText") && tipSource.helpText )
			registerCtrlHoverTip(ctrlObj, tipSource.helpText)
		return
	}

	registerCtrlHoverTip(ctrlObj, tipSource)
}

clearCtrlHoverTip()
{
	ToolTip()
}

; Show hover help for registered controls
; --------------------------------------
showCtrlHoverTip(wParam, lParam, msg, hwnd)
{
	global UI, gui1, APP_IS_CLOSING
	static lastCtrlHwnd := 0, tipVisible := false

	if ( APP_IS_CLOSING || !IsObject(gui1) ) {
		ToolTip()
		return
	}

	MouseGetPos(,, &winHwnd, &ctrlHwnd, 2)
	if ( winHwnd != gui1.Hwnd ) {
		if ( tipVisible )
			ToolTip()
		lastCtrlHwnd := 0
		tipVisible := false
		return
	}

	if ( ctrlHwnd == lastCtrlHwnd )
		return
	lastCtrlHwnd := ctrlHwnd

	if ( UI.HasOwnProp("ctrlHoverTips") && (UI.ctrlHoverTips is Map) && UI.ctrlHoverTips.Has(ctrlHwnd) ) {
		ToolTip(UI.ctrlHoverTips[ctrlHwnd])
		tipVisible := true
	}
	else {
		if ( tipVisible )
			ToolTip()
		tipVisible := false
	}
}


; Convert string to lowercase
; ----------------------------
stringLower(str) 
{
	return StrLower(str)
}


; Convert string to uppercase
; ---------------------------
stringUpper(str, title:=false) 
{
	return title ? StrTitle(str) : StrUpper(str)
}


; Check if value is in array
; ---------------------------
inArray(cVal, arry) 
{
	for idx, val in arry
		if ( cVal == val )
			return true
	return false
}


; Create a string of items seperated by a delimeter from an array
; ---------------------------------------------------------------
arrayToString(arr, delim:=", ")
{
	rtn := ""
	for idx, val in arr
		rtn .= val delim
	return rtn ? SubStr(rtn, 1, -StrLen(delim)) : ""
}


; Create a readable list string using "or" for the last item
; ----------------------------------------------------------
arrayToOrString(arr)
{
	if ( !IsObject(arr) || !arr.Length )
		return ""
	if ( arr.Length == 1 )
		return arr[1]
	if ( arr.Length == 2 )
		return arr[1] " or " arr[2]

	rtn := ""
	Loop arr.Length - 1
		rtn .= arr[A_Index] ", "
	return SubStr(rtn, 1, -2) ", or " arr[arr.Length]
}


; Create an Array from a string with delimeters
; ---------------------------------------------
arrayFromString(string, delim:=",") 
{
	rtnArray := []
	loop parse string, delim
		rtnArray.push(a_loopfield)
	return rtnArray
}


; Remove an item from an array
; -------------------------------
removeFromArray(removeItem, thisArray)
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

removeFinishedInputEntry(recvCmd, recvFromFileFull)
{
	global job, LV
	if ( !recvCmd || !recvFromFileFull || !job.scannedFiles.HasOwnProp(recvCmd) )
		return false

	targetNorm := StrLower(RegExReplace(recvFromFileFull, "/", "\"))
	removed := false

	for idx, val in job.scannedFiles.%recvCmd% {
		if ( StrLower(RegExReplace(val, "/", "\")) == targetNorm ) {
			job.scannedFiles.%recvCmd%.RemoveAt(idx)
			removed := true
			break
		}
	}

	if ( IsObject(LV) ) {
		loop LV.GetCount() {
			rtn := LV.GetText(A_Index)
			if ( StrLower(RegExReplace(rtn, "/", "\")) == targetNorm ) {
				LV.Delete(A_Index)
				removed := true
				break
			}
		}
	}
	return removed
}

resolveNearestExistingDir(dirPath, fallbackDir:="")
{
	defaultDir := InStr(FileExist(A_Desktop), "D") ? A_Desktop : A_WorkingDir
	checkDir := Trim(StrReplace(dirPath, "/", "\"), " `t`r`n`"")
	fallbackDir := Trim(StrReplace(fallbackDir, "/", "\"), " `t`r`n`"")

	for _, thisDir in [checkDir, fallbackDir, defaultDir] {
		currDir := thisDir
		while ( currDir ) {
			currDir := RegExReplace(currDir, "\\+$")
			if ( RegExMatch(currDir, "^[A-Za-z]:$") )
				currDir .= "\"
			if ( InStr(FileExist(currDir), "D") )
				return currDir
			parentDir := splitFilePath(currDir).dir
			if ( !parentDir || parentDir == currDir )
				break
			currDir := parentDir
		}
	}
	return defaultDir
}
; GUI state and control helpers
; ------------------------

; Get Windows current default font
; by SKAN
; ------------------------
guiDefaultFont() 
{
	LF := Buffer(92, 0) ; LOGFONT structure
	if DllCall("GetObject", "Ptr", DllCall("GetStockObject", "Int", 17, "Ptr"), "Int", LF.Size, "Ptr", LF.Ptr)
	return {name: StrGet(LF.Ptr + 28, 32), size: Round(Abs(NumGet(LF, 0, "Int")) * (72 / A_ScreenDPI), 1)
			, weight: NumGet(LF, 16, "Int"), quality: NumGet(LF, 26, "UChar")}
	return False
}


; GUI window was moved
; --------------------
moveGUIWin(wParam, lParam, msg, hwnd)
{
	global APP_MAIN_WIN_POS_X, APP_MAIN_WIN_POS_Y, APP_VERBOSE_WIN_POS_X, APP_VERBOSE_WIN_POS_Y, APP_MAIN_NAME, APP_VERBOSE_NAME, gui1, gui2, APP_IS_CLOSING

	if ( APP_IS_CLOSING )
		return
	
	if ( gui1 && hwnd == gui1.Hwnd ) {
		WinGetPos(&APP_MAIN_WIN_POS_X, &APP_MAIN_WIN_POS_Y, , , APP_MAIN_NAME)
		SetTimer(writeMoveGUIWin, -500)
	}
	else if ( gui2 && hwnd == gui2.Hwnd ) {
		WinGetPos(&APP_VERBOSE_WIN_POS_X, &APP_VERBOSE_WIN_POS_Y, , , APP_VERBOSE_NAME)
		SetTimer(writeMoveGUIWin, -500)
	}

	writeMoveGUIWin(*)
	{
		ini("write", ["APP_MAIN_WIN_POS_X", "APP_MAIN_WIN_POS_Y", "APP_VERBOSE_WIN_POS_X", "APP_VERBOSE_WIN_POS_Y"])
	}
	return
}


; Verbose window was resized
; --------------------------
GuiSize2(guiHwnd, eventInfo, W, H) 
{
	global APP_VERBOSE_WIN_HEIGHT := H, APP_VERBOSE_WIN_WIDTH := W
	
	SetTimer(write2GuiSize, -500)					

	write2GuiSize(*)
	{
		ini("write", ["APP_VERBOSE_WIN_HEIGHT", "APP_VERBOSE_WIN_WIDTH"])
	}
	return	
}
; Function replacement for guiControl
; ------------------------------------
; Example usages:
; guiCtrl({thisButton:"New Button Text"}, 1) - works on GUI #1
; guiCtrl("move", {stuff:"x9 w200", thing:"x1 y2"}, 3)

guiCtrl(arg1:="", arg2:="", arg3:="") 
{
	if ( isObject(arg1) )
		obj := arg1, guiNum := arg2 ? arg2 : 1, cmd := ""
	else
		obj := arg2, cmd := stringLower(arg1), guiNum := arg3 ? arg3 : 1

	for ele, newVal in obj.OwnProps() {
		ctrl := getGuiCtrl(ele, guiNum)
		if ( !ctrl )
			continue

		switch cmd {
			case "move", "movedraw":
				ctrl.GetPos(&x, &y, &w, &h)
				if ( RegExMatch(newVal, "i)(?:^|\s)x(-?\d+)", &m) )
					x := m[1]
				if ( RegExMatch(newVal, "i)(?:^|\s)y(-?\d+)", &m) )
					y := m[1]
				if ( RegExMatch(newVal, "i)(?:^|\s)w(-?\d+)", &m) )
					w := m[1]
				if ( RegExMatch(newVal, "i)(?:^|\s)h(-?\d+)", &m) )
					h := m[1]
				ctrl.Move(x, y, w, h)
				if ( cmd == "movedraw" )
					ctrl.Redraw()

			case "choose":
				ctrl.Choose(Trim(newVal, "|") + 0)

			case "opt":
				ctrl.Opt(newVal)

			case "":
				if ( (ctrl.Type == "DropDownList" || ctrl.Type == "DDL") && InStr(newVal, "|") ) {
					items := []
					chooseIdx := 0
					for idx, item in StrSplit(newVal, "|") {
						if ( item == "" ) {
							if ( items.Length && !chooseIdx )
								chooseIdx := items.Length
							continue
						}
						items.Push(item)
					}
					ctrl.Delete()
					if ( items.Length )
						ctrl.Add(items)
					if ( chooseIdx )
						ctrl.Choose(chooseIdx)
				}
				else if ( ctrl.Type == "Edit" || ctrl.Type == "Progress" )
					ctrl.Value := newVal
				else if ( ctrl.Type == "Pic" || ctrl.Type == "Picture" )
					ctrl.Value := newVal
				else if ( ctrl.Type == "CheckBox" || ctrl.Type == "Radio" ) {
					if ( RegExMatch(newVal "", "^-?\d+$") )
						ctrl.Value := newVal
					else
						ctrl.Text := newVal
				}
				else
					ctrl.Text := newVal
		}
	}
} 


; guiToggle GUI controls
; ------------------------------------------
guiToggle(doWhat, whichControls, guiNum:=1) 
{
	global gui1, gui2, gui3, gui4, gui5, APP_MAIN_NAME
	
	if ( !doWhat || !whichControls )
		return false
	
	doWhatArray := isObject(doWhat) ? doWhat : [doWhat]
	ctlArray := isObject(whichControls) ? whichControls : (whichControls == "all" ? getWinControls(APP_MAIN_NAME, "Static") : [whichControls])

	for idx, dw in doWhatArray
		for idx2, ctl in ctlArray {
			ctrlObj := getGuiCtrl(ctl, guiNum)
			if ( !ctrlObj )
				continue
			switch stringLower(dw) {
				case "enable":
					ctrlObj.Enabled := true
				case "disable":
					ctrlObj.Enabled := false
				case "show":
					ctrlObj.Visible := true
				case "hide":
					ctrlObj.Visible := false
			}
		}
			
}



; Function replacement for guiControlGet
; --------------------------------------
guiCtrlGet(ctrl, guiNum:=1) 
{
	ctrlObj := getGuiCtrl(ctrl, guiNum)
	if ( !ctrlObj )
		return ""
	if ( ctrlObj.Type == "Text" || ctrlObj.Type == "Button" || ctrlObj.Type == "GroupBox" || ctrlObj.Type == "Link" || ctrlObj.Type == "DropDownList" || ctrlObj.Type == "DDL" || ctrlObj.Type == "ComboBox" )
		return ctrlObj.Text
	return ctrlObj.Value
}



; Get a windows control elements as an array
; -------------------------------------------
getWinControls(win, ignoreStr:="") 
{
	rtnArray := []
	for idx, thisCtrl in WinGetControls(win)
	{
		if ( ignoreStr && inStr(thisCtrl, ignoreStr) ) ; Dont disable text elements
			continue
		rtnArray.push(thisCtrl)	
	}
	return rtnArray
}

showAutoCloseConfirm(msg, title:="", expireFn:="")
{
	global gui1, gui6, APP_MAIN_NAME
	result := ""

	if ( IsObject(expireFn) ) {
		try {
			if ( expireFn.Call() )
				return "AutoClose"
		}
	}

	if ( IsObject(gui6) )
		try gui6.Destroy()

	gui6 := Gui("+Owner" gui1.Hwnd " +AlwaysOnTop +ToolWindow", title ? title : APP_MAIN_NAME)
	gui6.MarginX := 16
	gui6.MarginY := 16
	gui6.SetFont("s9 Q5 w400 c000000")
	gui6.Add("text",, msg)
	btnYes := gui6.Add("button", "xm y+15 w95 Default", "Yes")
	btnNo := gui6.Add("button", "x+10 w95", "No")

	setResult(val, *)
	{
		result := val
	}

	checkExpire(*)
	{
		if ( result || !IsObject(expireFn) )
			return
		try {
			if ( expireFn.Call() )
				result := "AutoClose"
		}
	}

	btnYes.OnEvent("Click", setResult.Bind("Yes"))
	btnNo.OnEvent("Click", setResult.Bind("No"))
	gui6.OnEvent("Close", setResult.Bind("No"))
	gui6.Show("AutoSize Center")
	SetTimer(checkExpire, 100)

	while ( !result )
		Sleep(50)

	SetTimer(checkExpire, 0)
	if ( IsObject(gui6) )
		try gui6.Destroy()
	gui6 := ""
	return result
}
; Text formatting helpers
; --------------------
; Draw a line by count
; --------------------
drawLine(num:=1) 
{
	if ( num < 1 )
		return ""
	loop num
		rtn .= "-"
	return rtn	
}

; Delete a file 
; -------------
deleteFileWithRetry(file, attempts:=5, sleepdelay:=50) 
{
	f := fileExist(file)
	if ( !f || f == "D" ) 									; if file dosent exist or file is a directory 
		return true
	
	loop (attempts < 1 ? 1 : attempts) { 					; 5 attempts to delete the file
		try FileSetAttrib("-RSH", file)						; Clear common blocking attributes before deleting
		try FileDelete(file)
		catch
			; Retry below after short sleep.
		Sleep(sleepdelay)
		f := fileExist(file)
		if ( !f || f == "D" )							; Success
			return true
	}
	return false
}


; Delete a folder
; ---------------
folderDelete(dir, attempts:=5, sleepdelay:=50, full:=0) 
{
	if ( fileExist(dir) != "D" )								; If supplied dir isn't a directory, we are good to go
		return true
	;if ( dllCall("Shlwapi\PathIsDirectoryEmpty", "Str", dir) ) ; if empty
	loop (attempts < 1 ? 1 : attempts) {
		if ( full ) {
			loop Files RegExReplace(dir, "\\$") "\*.*", "FR" {
				try FileSetAttrib("-RSH", A_LoopFileFullPath)
			}
			loop Files RegExReplace(dir, "\\$") "\*.*", "DR" {
				try FileSetAttrib("-RSH", A_LoopFileFullPath)
			}
		}
		try FileSetAttrib("-RSH", dir)
		try DirDelete(dir, full)								; Attempt to delete the directory x times, full flag for when folder is full
		catch
			; Retry below after short sleep.
		Sleep(sleepdelay)	
		if ( fileExist(dir) != "D" )							; Success
			return true
	}
	return false
}

; Delete a list of files that have been successfuly deleted
deleteFilesReturnList(file) 
{
	delFiles := ""
	for idx, thisFile in getFilesFromCUEGDITOC(file)
		delFiles .= (deleteFileWithRetry(thisFile, 3, 100) ? thisFile ", " : "")
	return delFiles
}


; List filenames from CUE, GDI and TOC files 
; ------------------------------------------
getFilesFromCUEGDITOC(inputFiles) 
{
	fileList := []
	if ( !isObject(inputFiles) )
		inputFiles := [inputFiles]
		
	for idx, thisFile in inputFiles {
		if ( !fileExist(thisFile) )
			continue
		
		fileList.push(thisFile) ; Always include the file being supplied
		f := splitFilePath(thisFile)
		fExt := StringLower(f.ext)
		loop Read, thisFile 
		{
			switch fExt {
				case "cue", "toc":
					file := ""
					if ( RegExMatch(a_loopReadLine, "i)^\s*FILE\s+`"([^`"]+)`"", &mQuoted) )
						file := mQuoted[1]
					else if ( RegExMatch(a_loopReadLine, "i)^\s*FILE\s+([^\s]+)", &mBare) )
						file := mBare[1]
					if ( file && fileExist(f.dir "\" file) )
						fileList.push(f.dir "\" file)
				case "gdi":
					if ( RegExMatch(a_loopReadLine, "^\s*\d") && a_index > 1 ) {
						loop parse a_loopReadLine, " `;"
							if ( fileExist(f.dir "\" a_loopField) )
								fileList.push(f.dir "\" a_loopField)
					}
			}
		}
	}
	return fileList
}

detect7ZipExe()
{
	global SEVENZIP_CUSTOM_PATH
	pathEnv := EnvGet("PATH")
	customExe := resolve7ZipPath(SEVENZIP_CUSTOM_PATH)
	if ( customExe )
		return customExe
	for _, thisFile in ["C:\Program Files\7-Zip\7z.exe", "C:\Program Files (x86)\7-Zip\7z.exe"] {
		if ( FileExist(thisFile) )
			return thisFile
	}
	Loop Parse pathEnv, ";"
	{
		thisDir := Trim(A_LoopField, " `t`r`n`"")
		if ( !thisDir )
			continue
		thisFile := RegExReplace(thisDir, "\\+$") "\7z.exe"
		if ( FileExist(thisFile) )
			return thisFile
	}
	return ""
}

resolve7ZipPath(pathStr)
{
	if ( !pathStr )
		return ""
	checkPath := Trim(StrReplace(pathStr, "/", "\"), " `t`r`n`"")
	if ( !checkPath )
		return ""
	if ( InStr(FileExist(checkPath), "D") )
		checkPath := RegExReplace(checkPath, "\\+$") "\7z.exe"
	if ( !FileExist(checkPath) )
		return ""
	fileParts := splitFilePath(checkPath)
	return StrLower(fileParts.file) == "7z.exe" ? checkPath : ""
}

get7ZipBrowseDir()
{
	global SEVENZIP_CUSTOM_PATH, SEVENZIP_EXE
	for _, thisPath in [SEVENZIP_CUSTOM_PATH, SEVENZIP_EXE, "C:\Program Files\7-Zip", "C:\Program Files (x86)\7-Zip"] {
		if ( !thisPath )
			continue
		if ( InStr(FileExist(thisPath), "D") )
			return thisPath
		parentDir := splitFilePath(thisPath).dir
		if ( InStr(FileExist(parentDir), "D") )
			return parentDir
	}
	return resolveNearestExistingDir(A_Desktop, A_WorkingDir)
}

isArchiveExt(fileExt)
{
	ext := StrLower(Trim(fileExt, ". "))
	return ext == "zip" || ext == "7z"
}

; Read an archive file
; --------------------
readArchiveFile(archiveFile)
{
	ext := StrLower(splitFilePath(archiveFile).ext)
	if ( ext == "zip" )
		return readZipFile(archiveFile)
	if ( ext == "7z" )
		return read7ZipFile(archiveFile)
	return []
}

; Read a zip file
; ----------------
readZipFile(zipFile) 
{
	if ( splitFilePath(zipFile).ext != "zip" )
		return []
	try zippedItem := ComObject("Shell.Application").Namespace(zipFile)
	catch
		return []
	if ( !IsObject(zippedItem) )
		return []
	return readZipFolderItems(zippedItem.Items)
}

readZipFolderItems(folderItems, prefix:="")
{
	array := []
	for item in folderItems {
		itemPath := prefix ? prefix "\" item.Name : item.Name
		isFolder := false
		try isFolder := item.IsFolder
		if ( isFolder ) {
			try subFolder := item.GetFolder
			if ( IsObject(subFolder) ) {
				for _, subItemPath in readZipFolderItems(subFolder.Items, itemPath)
					array.Push(subItemPath)
			}
		}
		else
			array.Push(itemPath)
	}
	return array
}

read7ZipFile(archiveFile)
{
	global SEVENZIP_EXE
	q := Chr(34)
	tmpOut := A_Temp "\namdhc_7z_list_" A_TickCount "_" Random(1000, 9999) ".txt"
	items := []
	sawDivider := false

	if ( !SEVENZIP_EXE || !FileExist(SEVENZIP_EXE) )
		return []

	cmd := A_ComSpec . " /C " . q . q . SEVENZIP_EXE . q . " l -slt -sccUTF-8 -- " . q . archiveFile . q . " > " . q . tmpOut . q . " 2>&1" . q
	try RunWait(cmd, "", "Hide", &exitCode)
	catch
		return []

	if ( !FileExist(tmpOut) )
		return []

	try rawList := FileRead(tmpOut, "UTF-8")
	catch {
		try rawList := FileRead(tmpOut)
	}
	try FileDelete(tmpOut)

	Loop Parse rawList, "`n", "`r"
	{
		line := Trim(A_LoopField, " `t")
		if ( line == "----------" ) {
			sawDivider := true
			continue
		}
		if ( !sawDivider )
			continue
		if ( RegExMatch(line, "^Path = (.+)$", &m) )
			items.Push(m[1])
	}
	return items
}



/*
Milliseconds to HH:MM:SS
Thanks Odlanir
https://www.autohotkey.com/boards/viewtopic.php?t=45476
*/
millisecToTime(msec) 
{
	secs := floor(mod((msec / 1000), 60))
	mins := floor(mod((msec / (1000 * 60)), 60) )
	hour := floor(mod((msec / (1000 * 60 * 60)), 24))
	return Format("{:02}:{:02}:{:02}", hour, mins, secs)
}


; Convert bytes to human readable format
; By SKAN on CT5H/D351 @ tiny.cc/formatbytes
; https://www.autohotkey.com/boards/viewtopic.php?f=6&t=3567&sid=baefa38e09754ae9bfe4445026c545c8&start=20
formatBytes(N) { 
	Return DllCall("Shlwapi\StrFormatByteSize64A", "Int64",N, "Str",Format("{:16}",N), "Int",16, "AStr")
}

formatBytesReport(N) {
	return formatBytes(N) " (" formatBytesIEC(N) ")"
}

formatBytesIEC(N) {
	units := ["B", "KiB", "MiB", "GiB", "TiB", "PiB"]
	val := N + 0.0
	idx := 1
	while ( Abs(val) >= 1024 && idx < units.Length ) {
		val /= 1024.0
		idx++
	}
	if ( idx = 1 )
		return Round(val) " " units[idx]

	if ( Abs(val) >= 100 )
		s := Format("{:.0f}", val)
	else if ( Abs(val) >= 10 )
		s := Format("{:.1f}", val)
	else
		s := Format("{:.2f}", val)
	s := RegExReplace(s, "\.?0+$")
	return s " " units[idx]
}

; Network, update and app lifecycle helpers
; -----------------------------------------

; URL download to a file var
; Thanks maestrith 
; https://www.autohotkey.com/board/topic/88685-download-a-url-to-a-variable/
URLDownloadToVar(url){
	try {
		hObject:=ComObject("WinHttp.WinHttpRequest.5.1")
		hObject.Open("GET",url)
		hObject.Send()
		return hObject.ResponseText
	}
}

; Normalize path for Windows
;-----------------------------
normalizePath(path) {
    cc := DllCall("GetFullPathName", "str", path, "uint", 0, "ptr", 0, "ptr", 0, "uint")
    buf := Buffer(cc*2, 0)
    DllCall("GetFullPathName", "str", path, "uint", cc, "ptr", buf.Ptr, "ptr", 0)
    return StrGet(buf)
}


; Check github for for newest assets
; ----------------------------------
checkForUpdates(arg1:="", userClick:=false) 
{
	global CURRENT_VERSION, APP_MAIN_NAME, gui4
	if ( IsObject(gui4) )
		gui4.Opt('+OwnDialogs')

	log("Checking for updates ... ")

	releasePageFallback := "https://github.com/umageddon/namDHC/releases/latest"
	GITHUB_REPO_URL := "https://api.github.com/repos/umageddon/namDHC/releases/latest"
	JSONStr := URLDownloadToVar(GITHUB_REPO_URL)
	if ( !JSONStr ) {
		log("Error updating: Update info invalid")
		if ( userClick )
			MsgBox("Update info invalid", "Error getting update info", 16)
		return
	}

	try obj := main_jsonToLegacyObj(jsongo.Parse(JSONStr))
	catch {
		log("Error updating: Update info invalid")
		if ( userClick )
			MsgBox("Update info invalid", "Error getting update info", 16)
		return
	}
	if ( !isObject(obj) ) {
		log("Error updating: Update info invalid")
		if ( userClick )
			MsgBox("Update info invalid", "Error getting update info", 16)
		return
	}

	apiMessage := obj.HasOwnProp("message") ? obj.message : ""
	tagName := obj.HasOwnProp("tag_name") ? obj.tag_name : ""
	releaseBody := obj.HasOwnProp("body") ? obj.body : ""
	releaseUrl := obj.HasOwnProp("html_url") ? obj.html_url : releasePageFallback

	if ( apiMessage && inStr(apiMessage, "limit exceeded") ) {
		log("Error updating: Github API limit exceeded")
		if ( userClick )
			MsgBox("Github API limit exceeded", "Error getting update info", 16)
		return
	}

	if ( !tagName ) {
		log("Error updating: Update info invalid")
		if ( userClick )
			MsgBox("Update info invalid", "Error getting update info", 16)
		return
	}

	newVersion := regExReplace(tagName, "namDHC|v| ", "")
	cmp := VerCompare(newVersion, CURRENT_VERSION)
	if ( cmp = 0 ) {
		log("No new updates found. You are running the latest version")
		if ( userClick )
			MsgBox("You are running the latest version", "No new updates found", 64)
		return
	}

	if ( cmp < 0 ) {
		log("Your version is newer then the latest release!  Current version: v" CURRENT_VERSION " - Latest version: v" newVersion)
		if ( userClick )
			MsgBox("Your version is newer then the latest release!`n`nCurrent version: v" CURRENT_VERSION " - Latest version: v" newVersion, "Info", 64)
		return
	}

	log("An update was found: v" newVersion)
	msg := "A new version of " APP_MAIN_NAME " is available!`n`nCurrent version: v" CURRENT_VERSION
	msg .= "`nLatest version: v" newVersion
	msg .= "`n`nChanges:`n" strReplace(releaseBody, "-", "    -")
	msg .= "`n`nFor safety, namDHC opens GitHub in your browser for manual download/install."
	msg .= "`nOpen release page now?"
	if ( MsgBox(msg, "Update available", 36) == "Yes" )
		Run(releaseUrl)
}




; Kill all namDHC process (including chdman.exe)
; -----------------------------------------------
killAllProcess() 
{
	global APP_MAIN_NAME, APP_VERBOSE_NAME, APP_RUN_JOB_NAME, APP_RUN_CONSOLE_NAME
	currPID := DllCall("GetCurrentProcessId")

	loop {
		if ( !ProcessClose("chdman.exe") )
			break
		Sleep(100)
	}

	for idx, app in [APP_MAIN_NAME, APP_VERBOSE_NAME, APP_RUN_JOB_NAME, APP_RUN_CONSOLE_NAME] {
		hwnd := winExist(app)
		if ( !hwnd )
			continue
		if ( WinGetPID("ahk_id " hwnd) == currPID )
			continue
		WinActivate("ahk_id " hwnd)
		WinClose("ahk_id " hwnd)
		if ( winExist("ahk_id " hwnd) ) {
			PostMessage(0x0112, 0xF060, , , "ahk_id " hwnd)
			WinKill("ahk_id " hwnd)
		}
	}
}


quitApp() 
{
	global job, APP_MAIN_NAME, APP_IS_CLOSING

	APP_IS_CLOSING := true
	clearCtrlHoverTip()
	try OnMessage(0x200, showCtrlHoverTip, 0)
	try OnMessage(0x03, moveGUIWin, 0)

	if ( job.HasOwnProp("started") && job.started == true ) {
		if ( MsgBox("Are you sure? Currently running jobs will be killed.", APP_MAIN_NAME, 52) == "Yes" ) {
			killAllProcess()
			ExitApp()
		}
		return false
	}

	ExitApp()
}

#singleInstance off
#noEnv
#Persistent
detectHiddenWindows On
setTitleMatchmode 3
SetWorkingDir %a_ScriptDir%


#Include SelectFolderEx.ahk
#Include ClassImageButton.ahk
#Include ConsoleClass.ahk
#Include JSON.ahk


VER_HISTORY =
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
- Workaround to at least get the options crammed in view - for those folks who use higher DPI settings
 Thanks TFWol

v1.13
)


; Default global values 
; ---------------------
CURRENT_VERSION := "1.13"
CHECK_FOR_UPDATES_STARTUP := "yes"
CHDMAN_FILE_LOC := a_scriptDir "\chdman.exe"
DIR_TEMP := a_Temp "\namDHC"
CHDMAN_VERSION_ARRAY := ["0.239", "0.240", "0.241", "0.249"]
GITHUB_REPO_URL := "https://api.github.com/repos/umageddon/namDHC/releases/latest" 
APP_MAIN_NAME := "namDHC"
APP_VERBOSE_NAME := APP_MAIN_NAME " - Verbose"
APP_RUN_JOB_NAME := APP_MAIN_NAME " - Job"
APP_RUN_CHDMAN_NAME := APP_RUN_JOB_NAME " - chdman"
APP_RUN_CONSOLE_NAME := APP_RUN_JOB_NAME " - Console"
TIMEOUT_SEC := 20
WAIT_TIME_CONSOLE_SEC := 1
JOB_QUEUE_SIZE := 3
JOB_QUEUE_SIZE_LIMIT := 10
OUTPUT_FOLDER := a_workingDir
PLAY_SONG_FINISHED := "yes"
REMOVE_FILE_ENTRY_AFTER_FINISH := "yes"
SHOW_JOB_CONSOLE := "no"
SHOW_VERBOSE_WINDOW := "no"
APP_VERBOSE_WIN_HEIGHT := 400 
APP_VERBOSE_WIN_WIDTH := 800
APP_VERBOSE_WIN_POS_X := 775
APP_VERBOSE_WIN_POS_Y := 150
APP_MAIN_WIN_POS_X := 800
APP_MAIN_WIN_POS_Y := 100


; Read ini to write over globals if changed previously
;-------------------------------------------------------------
ini("read" 
	,["JOB_QUEUE_SIZE","OUTPUT_FOLDER","SHOW_JOB_CONSOLE","SHOW_VERBOSE_WINDOW","PLAY_SONG_FINISHED","REMOVE_FILE_ENTRY_AFTER_FINISH"
	,"APP_MAIN_WIN_POS_X","APP_MAIN_WIN_POS_Y","APP_VERBOSE_WIN_WIDTH","APP_VERBOSE_WIN_HEIGHT","APP_VERBOSE_WIN_POS_X","APP_VERBOSE_WIN_POS_Y","CHECK_FOR_UPDATES_STARTUP"])

if ( !fileExist(CHDMAN_FILE_LOC) ) {
	msgbox 16, % "Fatal Error", % "CHDMAN.EXE not found!`n`nMake sure the chdman executable is located in the same directory as namDHC and try again.`n`nThe following chdman verions are supported:`n" arrayToString(CHDMAN_VERSION_ARRAY)
	exitApp
}



; Run a chdman thread
; Will be called when running chdman - As to allow for a one file executable
;-------------------------------------------------------------
if ( a_args[1] == "threadMode" ) {
	#include threads.ahk
}


; Kill all processes so only one instance is running
;-------------------------------------------------------------
killAllProcess()


; Set working job variables
;-------------------------------------------------------------
job := {workTally:{}, msgData:[], availPSlots:[], workQueue:[], scannedFiles:{}, queuedMsgData:[], InputExtTypes:[], OutputExtType:[], selectedOutputExtTypes:[], selectedInputExtTypes:[]}

; Set GUI variables
;-------------------------------------------------------------
GUI := { chdmanOpt:{}, dropdowns:{job:{}, media:{}}, buttons:{normal:[], hover:[], clicked:[], disabled:[]}, menu:{namesOrder:[], File:[], Settings:[], About:[]} }
GUI.dropdowns.job := { 	 create: {pos:1,desc:"Create CHD files from media"}
						,extract: {pos:2,desc:"Extract images from CHD files"}
						,info: {pos:3, desc:"Get info from CHD files"}
						,verify: {pos:4, desc:"Verify CHD files"}}
						/*
						,addMeta: {pos:5, desc:"Add metadata to CHD files"}
						,delMeta: {pos:6, desc:"Delete metadata from CHD files"} 
						*/
GUI.dropdowns.media :=	{ cd:"CD image", hd:"Hard disk image", ld:"LaserDisc image", raw:"Raw image" }

GUI.buttons.default :=	{normal:[0, 0xFFCCCCCC, "", "", 3], hover:[0, 0xFFBBBBBB, "", 0xFF555555, 3], clicked:[0, 0xFFCFCFCF, "", 0xFFAAAAAA, 3], disabled:[0, 0xFFE0E0E0, "", 0xFFAAAAAA, 3] }
GUI.buttons.cancel :=	{normal:[0, 0xFFFC6D62, "", "White", 3], hover:[0, 0xFFff8e85, "", "White", 3], clicked:[0, 0xFFfad5d2, "", "White", 3], disabled:[0, 0xFFfad5d2, "", "White", 3]}
GUI.buttons.start :=	{normal:[0, 0xFF74b6cc, "", 0xFF444444, 3],	hover:[0, 0xFF84bed1, "", "White", 3], clicked:[0, 0xFFa5d6e6, "", "White", 3], disabled:[0, 0xFFd3dde0, "", 0xFF888888, 3] }	

; Set menu variables
;-------------------------------------------------------------
GUI.menu["namesOrder"] := ["File", "Settings", "About"]
GUI.menu.File[1] :=		{name:"Quit",											gotolabel:"quitApp",					saveVar:""}
GUI.menu.About[1] :=	{name:"About",											gotolabel:"menuSelected",				saveVar:""}
GUI.menu.Settings[1] :=	{name:"Check for updates automatically",				gotolabel:"menuSelected",				saveVar:"CHECK_FOR_UPDATES_STARTUP"}
GUI.menu.Settings[2] :=	{name:"Number of jobs to run concurrently",				gotolabel:":SubSettingsConcurrently",	saveVar:""}
GUI.menu.Settings[3] :=	{name:"Show a verbose window",							gotolabel:"menuSelected",				saveVar:"SHOW_VERBOSE_WINDOW", Fn:"showVerbose"}
GUI.menu.Settings[4] :=	{name:"Show a console window for each new job",			gotolabel:"menuSelected",				saveVar:"SHOW_JOB_CONSOLE"}
GUI.menu.Settings[5] :=	{name:"Play a sound when finished jobs",				gotolabel:"menuSelected",				saveVar:"PLAY_SONG_FINISHED"}
GUI.menu.Settings[6] :=	{name:"Remove file entry from list on success",			gotolabel:"menuSelected",				saveVar:"REMOVE_FILE_ENTRY_AFTER_FINISH"}

; misc GUI variables
;-------------------------------------------------------------
GUI.HDtemplate := { ddList: ""		; Hard drive template dropdown list
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
GUI.CPUCores := procCountDDList()

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
*/	

GUI.chdmanOpt.force :=				{name: "force",				paramString: "f",	description: "Force overwriting an existing output file"}
GUI.chdmanOpt.verbose :=			{name: "verbose",			paramString: "v",	description: "Verbose output", 								hidden: true}
GUI.chdmanOpt.outputBin :=			{name: "outputbin",			paramString: "ob",	description: "Output filename for binary data", 			editField: "filename.bin", useQuotes:true}
GUI.chdmanOpt.inputParent :=		{name: "inputparent", 		paramString: "ip",	description: "Input Parent", 								editField: "filename.ext", useQuotes:true}
GUI.chdmanOpt.inputStartFrame :=	{name: "inputstartframe", 	paramString: "isf",	description: "Input Start Frame", 							editField: 0}
GUI.chdmanOpt.inputFrames :=		{name: "inputframes", 		paramString: "if",	description: "Effective length of input in frames", 		editField: 0}
GUI.chdmanOpt.inputStartByte :=		{name: "inputstartbyte", 	paramString: "isb",	description: "Starting byte offset within the input", 		editField: 0}
GUI.chdmanOpt.outputParent :=		{name: "outputparent",		paramString: "op",	description: "Output parent file for CHD", 					editField: "filename.chd", useQuotes:true}
GUI.chdmanOpt.hunkSize :=			{name: "hunksize",			paramString: "hs",	description: "Size of each hunk (in bytes)", 				editField: 19584}
GUI.chdmanOpt.inputStartHunk :=		{name: "inputstarthunk",	paramString: "ish",	description: "Starting hunk offset within the input", 		editField: 0}
GUI.chdmanOpt.inputBytes :=			{name: "inputBytes",		paramString: "ib",	description: "Effective length of input (in bytes)", 		editField: 0}
GUI.chdmanOpt.compression :=		{name: "compression",		paramString: "c",	description: "Compression codecs to use", 					editField: "cdlz,cdzl,cdfl"}
GUI.chdmanOpt.inputHunks :=			{name: "inputhunks",		paramString: "ih",	description: "Effective length of input (in hunks)", 		editField: 0}
GUI.chdmanOpt.numProcessors :=		{name :"numprocessors",		paramString: "np",	description: "Max number of CPU threads to use", 			dropdownOptions: GUI.CPUCores}
GUI.chdmanOpt.template :=			{name: "template",			paramString: "tp",	description: "Hard drive template to use", 					dropdownOptions: GUI.HDtemplate.ddList, dropdownValues:GUI.HDtemplate.values}
GUI.chdmanOpt.chs :=				{name: "chs",				paramString: "chs",	description: "CHS Values [cyl, heads, sectors]", 			editField: "332,16,63"}
GUI.chdmanOpt.ident :=				{name: "ident",				paramString: "id",	description: "Name of ident file for CHS info", 			editField: "filename.chs", useQuotes:true}
GUI.chdmanOpt.size :=				{name: "size",				paramString: "s",	description: "Size of output file (in bytes)", 				editField: 0}
GUI.chdmanOpt.unitSize :=			{name: "unitsize",			paramString: "us",	description: "Size of each unit (in bytes)", 				editField: 0}
GUI.chdmanOpt.sectorSize :=			{name: "sectorsize",		paramString: "ss",	description: "Size of each hard disk sector (in bytes)", 	editField: 512}
GUI.chdmanOpt.deleteInputFiles :=	{name: "deleteInputFiles",						description: "Delete input files after completing job", 	masterOf: "deleteInputDir"}
GUI.chdmanOpt.deleteInputDir :=		{name: "deleteInputDir",						description: "Also delete input directory (if empty)", 				xInset:10}
GUI.chdmanOpt.createSubDir :=		{name: "createSubDir",							description: "Create a directory for each job"}
GUI.chdmanOpt.keepIncomplete :=		{name: "keepIncomplete",						description: "Keep failed or cancelled output files"}


; Create Main GUI and its elements
;-------------------------------------------------------------
createMainGUI()
createProgressBars() 
createMenus()

showVerbose(SHOW_VERBOSE_WINDOW)			; Check or uncheck item "Show verbose window"  and show the window 
selectJob()									; Select 1st selection in job dropdown list and trigger refreshGUI()

if ( CHECK_FOR_UPDATES_STARTUP == "yes" )
	checkForUpdates()

onMessage(0x03,		"moveGUIWin")			; If windows are moved, save positions in moveGUIWin()

sleep 25 									; Needed (?) to allow window to be detected
mainAppHWND := winExist(APP_MAIN_NAME)

log(APP_MAIN_NAME " ready.")

return

;-------------------------------------------------------------------------------------------------------------------------





; A Menu item was selected
;-------------------------------------------------------------
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
			loop % JOB_QUEUE_SIZE_LIMIT											; Uncheck all 
				menu, SubSettingsConcurrently, UnCheck, % a_index
			menu, SubSettingsConcurrently, Check, % A_ThisMenuItemPos			; Check selected
			JOB_QUEUE_SIZE := A_ThisMenuItemPos									; Set variable
			ini("write", "JOB_QUEUE_SIZE")
			log("Saved JOB_QUEUE_SIZE")
		
		case "AboutMenu":														; Menu: About
			guiToggle("disable", "all")
			gui 1:+OwnDialogs
			
			gui 4: destroy
			gui 4: margin, 20 20
			gui 4: font, s15 Q5 w700 c000000
			gui 4: add, text, x10 y10, % APP_MAIN_NAME
			gui 4: font, s10 Q5 w700 c000000
			gui 4: add, text, x100 y17, % " v" CURRENT_VERSION
			gui 4: font, s10 Q5 w400 c000000
			gui 4: add, text, x10 y35, % "A Windows frontend for the MAME CHDMAN tool"
			gui 4: add, button, x10 y65 w130 h22 gcheckForUpdates, % "Check for updates"

			gui 4: add, text, x10 y100, % "History"
			gui 4: add, edit, x10 y120 h200 w775, % VER_HISTORY
			gui 4: add, link, x10 y340, Github: <a href="https://github.com/umageddon/namDHC">https://github.com/umageddon/namDHC</a>
			gui 4: add, link, x10 y360, MAME Info: <a href="https://www.mamedev.org/">https://www.mamedev.org/</a>
			gui 4: font, s9 Q5 w400 c000000
			gui 4: add, text, x10 y390, % "(C) Copyright 2022 Umageddon"
			gui 4: show, w800 center, About
			Gui 4:+LastFound +AlwaysOnTop +ToolWindow
			controlFocus,, About 												; Removes outline around html anchor
			return
	}
}

4GuiClose() 
{
	gui 4:destroy
	refreshGUI()
	return
}


; Job selection
;-------------------------------------------------------------
selectJob() 
{
	global
	
	gui 1:submit, nohide
	gui 1:+ownDialogs
	
	; Changes depending on job selected
	switch dropdownJob {
		case GUI.dropdowns.job.create.desc:
			newStartButtonLabel := "CREATE CHD"	
			guiCtrl({dropdownMedia:"|" GUI.dropdowns.media.cd "|" GUI.dropdowns.media.hd "|" GUI.dropdowns.media.ld "|" GUI.dropdowns.media.raw})

		case GUI.dropdowns.job.extract.desc:											
			newStartButtonLabel := "EXTRACT MEDIA"
			guiCtrl({dropdownMedia:"|" GUI.dropdowns.media.cd "|" GUI.dropdowns.media.hd "|" GUI.dropdowns.media.ld "|" GUI.dropdowns.media.raw})

		case GUI.dropdowns.job.info.desc:
			newStartButtonLabel := "GET INFO"
			guiCtrl({dropdownMedia:"|CHD Files"})

		case GUI.dropdowns.job.verify.desc:
			guiToggle("enable", "all")
			newStartButtonLabel := "VERIFY CHD"
			guiCtrl({dropdownMedia:"|CHD Files"})
		
		case GUI.dropdowns.job.addMeta.desc:
			newStartButtonLabel := "ADD METADATA"
			guiCtrl({dropdownMedia:"|CHD Files"})
			msgbox 64, % "", % "Option not implemented yet"
		
		case GUI.dropdowns.job.delMeta.desc:
			newStartButtonLabel := "DELETE METADATA"
			guiCtrl({dropdownMedia:"|CHD Files"})
			msgbox 64, % "", % "Option not implemented yet"
	}
	
	guiCtrl({buttonStartJobs:newStartButtonLabel})																									; New start button label to reflect new job
	imageButton.create(startButtonHWND, GUI.buttons.start.normal, GUI.buttons.start.hover, GUI.buttons.start.clicked, GUI.buttons.start.disabled)	; Default button colors,  must be set after changing button text

	guiCtrl("choose", {dropdownMedia:"|1"}) 																										; Choose first item in media dropdown and fire the selection 
}


; Media selection
;-------------------------------------------------------------
selectMedia()
{
	global
	local mediaSel, key, opt, val, idx, optNum, checkOpt, changeOpt, changeOptVal, ctrlY, x, y, gPos, file
	local optPerSide:=9, ctrlH:=25
	
	gui 1:submit, nohide
	gui 1:+ownDialogs

	; User selected media
	switch dropdownMedia {
		case GUI.dropdowns.media.cd: 	mediaSel := "cd"
		case GUI.dropdowns.media.hd:	mediaSel := "hd"
		case GUI.dropdowns.media.ld:	mediaSel := "ld"
		case GUI.dropdowns.media.raw:	mediaSel := "raw"
		default: mediaSel := "chd"
	}
	
	; Assign job variables according to media
	switch dropdownJob {
		case GUI.dropdowns.job.create.desc:			job.Cmd := "create" mediaSel, 	job.Desc := "Create CHD from a " stringUpper(mediaSel) " image",	job.FinPreTxt := "Jobs created"
		case GUI.dropdowns.job.extract.desc:		job.Cmd := "extract" mediaSel,	job.Desc := "Extract a " stringUpper(mediaSel) " image from CHD",	job.FinPreTxt := "Jobs extracted"
		case GUI.dropdowns.job.info.desc:			job.Cmd := "info", 				job.Desc := "Get info from CHD",									job.FinPreTxt := "Read info from jobs"
		case GUI.dropdowns.job.verify.desc:			job.Cmd := "verify",			job.Desc := "Verify CHD",											job.FinPreTxt := "Jobs verified"
		case GUI.dropdowns.job.addMeta.desc:		job.Cmd := "addmeta", 			job.Desc := "Add Metadata to CHD",									job.FinPreTxt := "Jobs with metadata added to"
		case GUI.dropdowns.job.delMeta.desc:		job.Cmd := "delmeta", 			job.Desc := "Delete Metadata from CHD",								job.FinPreTxt := "Jobs with metadata deleted from"
	}
	
	; Assign rest of job variables according to job
	switch job.Cmd {
		case "extractcd":	job.InputExtTypes := ["chd"],								job.OutputExtTypes := ["cue", "toc", "gdi"],	job.Options := [GUI.chdmanOpt.force, GUI.chdmanOpt.createSubDir, GUI.chdmanOpt.deleteInputFiles, GUI.chdmanOpt.keepIncomplete, GUI.chdmanOpt.outputBin, GUI.chdmanOpt.inputParent]
		case "extractld":	job.InputExtTypes := ["chd"],								job.OutputExtTypes := ["raw"],					job.Options := [GUI.chdmanOpt.force, GUI.chdmanOpt.createSubDir, GUI.chdmanOpt.deleteInputFiles, GUI.chdmanOpt.keepIncomplete, GUI.chdmanOpt.inputParent, GUI.chdmanOpt.inputStartFrame, GUI.chdmanOpt.inputFrames]
		case "extracthd":	job.InputExtTypes := ["chd"],								job.OutputExtTypes := ["img"],					job.Options := [GUI.chdmanOpt.force, GUI.chdmanOpt.createSubDir, GUI.chdmanOpt.deleteInputFiles, GUI.chdmanOpt.keepIncomplete, GUI.chdmanOpt.inputParent, GUI.chdmanOpt.inputStartByte, GUI.chdmanOpt.inputStartHunk, GUI.chdmanOpt.inputBytes, GUI.chdmanOpt.inputHunks]
		case "extractraw":	job.InputExtTypes := ["chd"],								job.OutputExtTypes := ["img", "raw"],			job.Options := [GUI.chdmanOpt.force, GUI.chdmanOpt.createSubDir, GUI.chdmanOpt.deleteInputFiles, GUI.chdmanOpt.keepIncomplete, GUI.chdmanOpt.inputParent, GUI.chdmanOpt.inputStartByte, GUI.chdmanOpt.inputStartHunk, GUI.chdmanOpt.inputBytes, GUI.chdmanOpt.inputHunks]
		case "createcd": 	job.InputExtTypes := ["cue", "toc", "gdi", "iso", "zip"],	job.OutputExtTypes := ["chd"],					job.Options := [GUI.chdmanOpt.force, GUI.chdmanOpt.createSubDir, GUI.chdmanOpt.deleteInputFiles, GUI.chdmanOpt.deleteInputDir, GUI.chdmanOpt.keepIncomplete, GUI.chdmanOpt.numProcessors, GUI.chdmanOpt.outputParent, GUI.chdmanOpt.hunkSize, GUI.chdmanOpt.compression]
		case "createld":	job.InputExtTypes := ["raw", "zip"],						job.OutputExtTypes := ["chd"],					job.Options := [GUI.chdmanOpt.force, GUI.chdmanOpt.createSubDir, GUI.chdmanOpt.deleteInputFiles, GUI.chdmanOpt.deleteInputDir, GUI.chdmanOpt.keepIncomplete, GUI.chdmanOpt.numProcessors, GUI.chdmanOpt.outputParent, GUI.chdmanOpt.inputStartFrame, GUI.chdmanOpt.inputFrames, GUI.chdmanOpt.hunkSize, GUI.chdmanOpt.compression]
		case "createhd":	job.InputExtTypes := ["img", "zip"],						job.OutputExtTypes := ["chd"],					job.Options := [GUI.chdmanOpt.force, GUI.chdmanOpt.createSubDir, GUI.chdmanOpt.deleteInputFiles, GUI.chdmanOpt.deleteInputDir, GUI.chdmanOpt.keepIncomplete, GUI.chdmanOpt.numProcessors, GUI.chdmanOpt.compression, GUI.chdmanOpt.outputParent, GUI.chdmanOpt.size, GUI.chdmanOpt.inputStartByte, GUI.chdmanOpt.inputStartHunk, GUI.chdmanOpt.inputBytes, GUI.chdmanOpt.inputHunks, GUI.chdmanOpt.hunkSize, GUI.chdmanOpt.ident, GUI.chdmanOpt.template, GUI.chdmanOpt.chs, GUI.chdmanOpt.sectorSize]
		case "createraw":	job.InputExtTypes := ["img", "raw", "zip"],					job.OutputExtTypes := ["chd"],					job.Options := [GUI.chdmanOpt.force, GUI.chdmanOpt.createSubDir, GUI.chdmanOpt.deleteInputFiles, GUI.chdmanOpt.deleteInputDir, GUI.chdmanOpt.keepIncomplete, GUI.chdmanOpt.numProcessors, GUI.chdmanOpt.outputParent, GUI.chdmanOpt.inputStartByte, GUI.chdmanOpt.inputStartHunk, GUI.chdmanOpt.inputBytes, GUI.chdmanOpt.inputHunks, GUI.chdmanOpt.hunkSize, GUI.chdmanOpt.unitSize, GUI.chdmanOpt.compression]
		case "info":		job.InputExtTypes := ["chd", "zip"],						job.OutputExtTypes := [],						job.Options := []
		case "verify":		job.InputExtTypes := ["chd", "zip"],						job.OutputExtTypes := [],						job.Options := []
		case "addmeta":		job.InputExtTypes := ["chd"],								job.OutputExtTypes := [],						job.Options := []
		case "delmeta":		job.InputExtTypes := ["chd"],								job.OutputExtTypes := [],						job.Options := []
	}	
	
	; Hide and uncheck ALL options
	for key, opt in GUI.chdmanOpt {
		guiToggle("hide", [opt.name "_checkbox", opt.name "_edit", opt.name "_dropdown"])	
		guiCtrl({(opt.name "_checkbox"):0})
	}
	
	; Show checkbox options depending on media selected
	if ( job.Options.length() )	{													
		guiToggle("show", "groupboxOptions")
		gPos := guiGetCtrlPos("groupboxOptions")														; Assign x & y values to groupbox x & y positions
		idx := 0, ctrlY := 0, x := gPos.x+10, y := gPos.y										
		
		for optNum, opt in job.Options {																; Show appropriate options according to selected type of job and media type
			if ( !opt || opt.hidden )
				continue
			
			if ( idx == optPerSide )																	; We've run out of vertical room to show more chdman options so new options go into next column
				x += 400, y := gPos.y, ctrlY := 1, idx := 0												; Set initial control/checkbox positions
			else 
				ctrlY++, idx++																			; Add to Y position for next control
			
			checkOpt := opt.name "_checkbox"
			guiCtrl({(checkOpt):" " (opt.description? opt.description : opt.name)}) 					; Label the checkbox -- Default to using the chdman opt.name parameter if no description
			guiCtrl("moveDraw", {(checkOpt):"x" x + (opt.xInset? opt.xInset:0) " y" y+(ctrlY*ctrlH)})	; Move the chdman checkbox 
			
			if ( opt.hasKey("editField") ) {															; The option can have either a dropdown or editfield
				changeOpt := opt.name "_edit"
				guiCtrl({(changeOpt):opt.editField})													; Populate editfield from GUI.chdmanOpt array
			}
			else if ( opt.hasKey("dropdownOptions") ) {
				changeOpt := opt.name "_dropdown"
				guiCtrl({(changeOpt):opt.dropdownOptions})												; Populate option dropdown from GUI.chdmanOpt array
			}
			guiCtrl("moveDraw", {(changeOpt):"x" x+210 " y" y +(ctrlY*ctrlH)-3})						; Move the option editfield or dropdown into place
			guiToggle("show", [checkOpt, changeOpt])													; Show the option and its coresponding editfield or dropdownlist
			guiCtrl({(checkOpt):opt.checked ? 1:0})														; Check it if by default it's on
		}
		guiCtrl("moveDraw", {groupboxOptions:"h" ceil((optNum >optPerSide ? optPerSide : optNum)*ctrlH)+30})	; Resize options groupbox height to fit all  options
	}
	else {																								; If no options to show for this selection ...
		guiToggle("hide", "groupboxOptions")															; ... then hide the groupbox
		guiCtrl("moveDraw", {groupboxOptions:"h0"})														; and resize the groupbox height for no options
	}
	
	; Reset extension menus
	menuExtHandler(true) 																

	; Populate listview
	gui 1: listView, listViewInputFiles						
	LV_delete()																							; Delete all listview entries
	for idx, file in job.scannedFiles[job.Cmd]
		LV_add("", file)																				; Re-populate listview with scanned files
	controlFocus, SysListView321, % APP_MAIN_NAME															; Focus on listview  to stop one item being selected
	
	refreshGUI()
}


; Refreshes GUI to reflect current settings
;-------------------------------------------------------------
refreshGUI() 
{
	global
	local opt, key, optNum
	static selectedJob
	
	gui 1:submit, nohide
	
	; By default, enable all elements
	guiToggle("enable", "all")
	
	; Show the main menu
	toggleMainMenu("show")
	
	; Changes to elements depending on job selected
	switch dropdownJob {
		case GUI.dropdowns.job.create.desc:
			
		case GUI.dropdowns.job.extract.desc:											
			
		case GUI.dropdowns.job.info.desc:
			guiToggle("disable", ["dropdownMedia", "buttonOutputExtType", "buttonInputExtType", "editOutputFolder", "buttonBrowseOutput"])

		case GUI.dropdowns.job.verify.desc:
			guiToggle("disable", ["dropdownMedia", "buttonOutputExtType", "buttonInputExtType", "editOutputFolder", "buttonBrowseOutput", "buttonBrowseOutput"])
			
		case GUI.dropdowns.job.addMeta.desc:
			guiToggle("disable", "all")
			guiToggle("enable", "dropdownJob")
		
		case GUI.dropdowns.job.delMeta.desc:
			guiToggle("disable", "all")
			guiToggle("enable", "dropdownJob")
	}
	
	; Enable chdman option checkboxes depending on job selected
	for optNum, opt in job.Options															
		guiToggle("enable", opt.name "_checkbox")
	

	; Checked option: enable or disable chdman option editfields, dropdowns or slave options 
	for optNum, opt in job.Options														
		guiToggle((guiCtrlGet(opt.name "_checkbox") ? "enable":"disable"), [opt.name "_dropdown", opt.name "_edit", opt.masterOf "_checkbox", opt.masterOf "_dropdown", opt.masterOf "_edit"]) ; If checked, enable or disable dropdown or editfields

	
	; Hide progress bars, progress text & progress groupbox
	guiToggle("hide", ["progressAll", "progressTextAll", "groupBoxProgress"])
	loop % JOB_QUEUE_SIZE_LIMIT
		guiToggle("hide", ["progress" a_index, "progressText" a_index, "progressCancelButton" a_index])


	; Trigger listview() to refresh selection buttons
	listViewInputFiles()
	
	; Set status text
	SB_SetText(job.scannedFiles[job.Cmd].length()? job.scannedFiles[job.Cmd].length() " jobs in the " stringUpper(job.Cmd) " queue. Ready to start!" : "Add files to the job queue to start", 1)
	
	; Make sure main button is showing
	guiToggle("hide", "buttonCancelAllJobs")
	guiToggle("show", "buttonStartJobs")
	guiToggle((job.scannedFiles[job.Cmd].length()>0 ? "enable":"disable"), ["buttonStartJobs", "buttonselectAllInputFiles"]) 		; Enable start button if there are jobs in the listview
	
	; Show and resize main GUI
	gui 1:show, autosize x%APP_MAIN_WIN_POS_X% y%APP_MAIN_WIN_POS_Y%, % APP_MAIN_NAME
	
	; Select main listview
	gui 1: listView, listViewInputFiles
}


; User pressed input or output files button
; Show Ext menu
;-------------------------------------------------------------
buttonExtSelect()
{
	switch a_guicontrol {
		case "buttonInputExtType":
			menu, InputExtTypes, Show, 873, 172 					; Hardcoded because x,y position returns from function are wonky
		case "buttonOutputExtType":
			menu, OutputExtTypes, Show, 873, 480
	}
}

	
; User selected an extension from the input/output extension menu
;-------------------------------------------------------------
menuExtHandler(init:=false)
{
	global job
	
	if ( init == true ) {													; Create and populate extension menu lists
		for idx, type in ["InputExtTypes", "OutputExtTypes"] {	
			menu, % type, deleteAll											; Clear all old Input & Output menu items
			
			job["selected" type] := []										; Clear global Array of selected Input & Output extensions
			for idx2, ext in job[type] {									; Parse through job.InputExtTypes & job.OutputExtTypes
				if ( !ext )
					continue
				menu, % type, add, % ext, % "menuExtHandler"				; Add extension item to the menu
				if ( dropdownJob == GUI.dropdowns.job.extract.desc && type == "OutputExtTypes" && idx2 > 1 )
					continue												; By default, only check one extension of the Output menu if we are extracting an image
				else {
					menu, % type, Check, % ext								; Otherwise, check all extension menu items
					job["selected" type].push(ext)							; Then add it to the input & output global selected extension array
				}
			}
		}
	}
	else if ( a_ThisMenu ) {												; An extension menu was selected
		selectedExtList := "selected" strReplace(a_ThisMenu, "extTypes", "") "ExtTypes"
		job[selectedExtList] := []											; Re-build either of these Arrays: job.selectedOutputExtTypes[] or job.selectedInputExtTypes[]
		
		switch a_ThisMenu {													; a_ThisMenu is either 'InputExtTypes' or 'OutputExtTypes'
			case "OutputExtTypes":											; Only one output extension is allowed to be checked
				;for idx, val in job.OutputExtTypes
				;menu, OutputExtTypes, Uncheck, % val						; Uncheck all menu items, 
				;menu, OutputExtTypes, Check, % a_ThisMenuItem				; Then check what was clicked, so only one is ever checked
				menu, OutputExtTypes, Togglecheck, % a_ThisMenuItem			; Toggle checking item
		
			case "InputExtTypes": 
				menu, InputExtTypes, Togglecheck, % a_ThisMenuItem			; Toggle checking item
				
		}
		for idx, val in job[a_ThisMenu]
			if ( isMenuChecked(a_ThisMenu, idx) ) {
				job[selectedExtList].push(val)								; Add checked extension item(s) to the global array for reference later
			}
		if ( job[selectedExtList].length() == 0 ) {
			menu, % a_ThisMenu, check, % a_ThisMenuItem						; Make sure at least one item is checked
			job[selectedExtList].push(a_ThisMenuItem)
		}
	}
	
	for idx, type in ["InputExtTypes", "OutputExtTypes"]					; Redraw input & output extension lists
		guiCtrl({(type) "Text": arrayToString(job["selected" type])})
}


; Scan files and add to queue
;-------------------------------------------------------------
addFolderFiles()
{
	global job, APP_MAIN_NAME
	newFiles := [], extList := "", numAdded := 0
	
	gui 1:submit, nohide
	gui 1:+OwnDialogs 
	
	guiToggle("disable", "all")
	
	if ( !isObject(job.scannedFiles[job.Cmd]) )
		job.scannedFiles[job.Cmd] := []
	
	switch ( a_GuiControl ) {
		case "buttonAddFiles":
			for idx, ext in job.selectedInputExtTypes
				extList .= "*." ext "; "
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
			inputFolder := selectFolderEx("", "Select a folder containing " arrayToString(job.selectedInputExtTypes) " type files.", winExist(APP_MAIN_NAME))
			if ( inputFolder.SelectedDir ) {
				inputFolder := regExReplace(inputFolder.SelectedDir, "\\$")
				
				for idx, thisExt in job.selectedInputExtTypes {
					loop, Files, % inputFolder "\*." thisExt, FR 
						newFiles.push(a_LoopFileLongPath)
				}
			}
	}
	
	if ( newFiles.length() ) {
		for idx, thisFile in newFiles {
			addFile := true, msg := "", fileParts := splitPath(thisFile)
			
			if ( !inArray(fileParts.ext, job.selectedInputExtTypes) )
				addFile := false ;,msg := "Skip adding " thisFile "  -  Not a selected format"

			else if ( inArray(thisFile, job.scannedFiles[job.Cmd]) )
				addFile := false, msg := "Skip adding " thisFile "  -  Already in queue"

			else if ( fileParts.ext == "zip" ) { 																; If its a zipfile, check to see if user extensions are contained within it
				addFile := false, msg := "Skip adding " thisFile "  -  No selected formats found in zipfile" 	; By default we assume zipfile dosent cantain a selected format
				for idx, fileInZip in readZipFile(thisFile) {
					zipFileExt := splitPath(fileInZip).ext
					if ( inArray(zipFileExt, job.selectedInputExtTypes) && zipFileExt <> "zip" ) {				; Only add file inside the zipfile if it is in the selected extension list and it's not another zipfile within a zipfile
						addFile := true
						continue
					}
				}
			}

			if ( addFile ) {
				numAdded++
				msg := "Adding " (fileInZip ? thisFile "   -->   " fileInZip : thisFile)
				LV_add("", thisFile)
				job.scannedFiles[job.Cmd].push(thisFile)
			}
			
			log(msg)
			SB_SetText(msg, 1, 1)
		}
	}
	reportQueuedFiles()
	refreshGUI()
}


; Listview containting input files was clicked
;-------------------------------------------------------------
listViewInputFiles()
{
	global
	local suffx, idx, val
	
	if ( a_guievent == "I" )
		return
	if ( LV_GetCount("S") > 0 )  {
		suffx := (LV_GetCount("S") > 1) ? "s" : ""
		guiCtrl({buttonRemoveInputFiles:"Remove file" suffx, buttonclearInputFiles:"Clear selection" suffx})
		for idx, val in [4,6]																					; Apply buttons 'Remove selections' and 'Clear selections' skin after changing text
			imageButton.create(GUIbutton%val%, GUI.buttons.default.normal, GUI.buttons.default.hover, GUI.buttons.default.clicked, GUI.buttons.default.disabled) ; Default button colors
		guiToggle("enable", ["buttonRemoveInputFiles", "buttonclearInputFiles"])
	}
	else
		guiToggle("disable", ["buttonclearInputFiles", "buttonRemoveInputFiles"]) 
		
	guiToggle((LV_GetCount()>0?"enable":"disable"), ["buttonselectAllInputFiles"]) 	

}


; Select from input listview files
; --------------------------------
selectInputFiles()
{
	global job
	row := 0, removeThese := []
	
	gui 1:submit, nohide
	gui 1:+OwnDialogs
	gui 1: listView, listViewInputFiles 					; Select the listview for manipulation
	
	if ( inStr(a_GuiControl, "SelectAll") ) {
		LV_Modify(0, "Select")
	}
	else if ( inStr(a_GuiControl, "Clear") ) {
		LV_Modify(0, "-Select")
	}
	else if ( inStr(a_GuiControl, "Remove") ) {
		guiToggle("disable", "all")
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
			removeFromArray(removeThisFile, job.scannedFiles[job.Cmd])
			log("Removed '" removeThisFile "' from the " stringUpper(job.Cmd) " queue")
			SB_setText("Removed '" removeThisFile "' from the " stringUpper(job.Cmd) " queue", 1)
		}
		reportQueuedFiles()
	}
	refreshGUI()
	controlFocus, SysListView321, %APP_MAIN_NAME%
}

reportQueuedFiles() 
{
	global job
	log( job.scannedFiles[job.Cmd].length()? job.scannedFiles[job.Cmd].length() " jobs in the " stringUpper(job.Cmd) " queue. Ready to start!" : "No jobs in the " stringUpper(job.Cmd) " file queue" )
	SB_setText( job.scannedFiles[job.Cmd].length()? job.scannedFiles[job.Cmd].length() " jobs in the " stringUpper(job.Cmd) " queue. Ready to start!" : "No jobs in the " stringUpper(job.Cmd) " file queue" )
}


; Select output folder
; --------------------
editOutputFolder()
{
	setTimer checkNewOutputFolder, -500
	return
}

; Check new inputted folder
; -------------------------
checkNewOutputFolder()
{
	global editOutputFolder, OUTPUT_FOLDER
	gui 1:submit, nohide
	gui 1:+ownDialogs

	if ( a_guiControl == "buttonBrowseOutput" ) {
		selectFolder := selectFolderEx(OUTPUT_FOLDER, "Select a folder to save converted files to", mainAppHWND)
		guiCtrl({editOutputFolder:selectFolder.selectedDir}) ; Assign to edit field if user selected
	}

	newFolder := editOutputFolder
	
	if ( !newFolder || newFolder == OUTPUT_FOLDER ) 
		return

	badChar := false
	for idx, val in ["*", "?", "<", ">", "/", "|", """"] {
		if ( inStr(newfolder, val) ) {
			badChar := true
			break
		}
	}
	folderChk := splitPath(newFolder)
	if ( !folderChk.drv || !folderChk.dir || badChar || strlen(newFolder) > 255 ) 	{		; Make sure newFolder is a valid directory string
		msgBox % "Invalid output folder"
		guiCtrl({editOutputFolder:OUTPUT_FOLDER})				; Edit reverts back to old value if new value invalid
	} else {
		OUTPUT_FOLDER := normalizePath(regExReplace(newFolder, "\\$"))

		log("'" OUTPUT_FOLDER "' selected as output folder")
		SB_SetText("'" OUTPUT_FOLDER "' selected as new output folder" , 1)
		ini("write", "OUTPUT_FOLDER")
		refreshGUI()
		controlFocus,, %APP_MAIN_NAME% 
	}
}



; A chdman option checkbox was clicked
; ---------------------------------
checkboxOption(ctrl:="")
{
	global GUI
	
	gui 1:submit, nohide
	opt := strReplace(a_guicontrol?a_guicontrol:ctrl, "_checkbox")
	guiToggle((%a_guiControl%? "enable":"disable"), [opt "_dropdown", opt "_edit"])		; Enable or disable corepsnding dropdown or editfield according to checked status
	if ( GUI.chdmanOpt[opt].hasKey("masterOf") ) {										; Disable the 'slave' checkbox if masterOf is set as an option
		guiToggle((%a_guiControl%? "enable":"disable"), [GUI.chdmanOpt[opt].masterOf "_checkbox", GUI.chdmanOpt[opt].masterOf "_dropdown", GUI.chdmanOpt[opt].masterOf "_edit"])
		guiCtrl({(GUI.chdmanOpt[opt].masterOf "_checkbox"):0})
	}
}



; Convert job button - Start the jobs!
; ------------------------------------
buttonStartJobs()
{
	global
	local fnMsg, runCmd, thisJob, gPos, y, file, filefull, dir, x1, x2, y, cmd, x, qsToDo
	static CHDInfoFileNum
	gui 1:submit, nohide
	gui 1:+ownDialogs
	
	switch job.Cmd { 
		case "createcd", "createhd", "createraw", "createld", "extractcd", "extracthd", "extractraw", "extractld":
			SB_SetText("Creating " stringUpper(job.Cmd) " work queue" , 1)
			log("Creating " stringUpper(job.Cmd) " work queue" )
			job.workQueue := createjob(job.Cmd, job.Options, job.selectedOutputExtTypes, job.selectedInputExtTypes, job.scannedFiles[job.Cmd])	; Create a queue (object) of files to process
			
		case "verify":
			SB_SetText("Verifying CHD's" , 1)
			log("Starting Verify CHD's" )
			job.workQueue := createjob("verify", "", "", job.selectedInputExtTypes, job.scannedFiles["verify"])
		
		case "info":
			SB_SetText("Info for CHD's" , 1)
			log("Getting info from CHD's" )
			guiToggle("disable", "all")
			CHDInfoFileNum := 1, x1 := 20, x2:= 150, y := 20
			gui 3: destroy
			gui 3: margin, 20, 20
			gui 3: font, s11 Q5 w750 c000000
			gui 3: add, text, x20 y%y% w700 vtextCHDInfoTitle, % ""
			y += 40
			loop 12 {
				gui 3: font, s9 Q5 w700 c000000
				gui 3: add, text, x%x1% y%y% w100 vtextCHDInfo_%a_index%, % ""
				gui 3: font, s9 Q5 w400 c000000
				gui 3: add, edit, x%x2% y+-13 w615 veditCHDInfo_%a_index% readonly, % ""
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
			y+=5
			gui 3: add, dropdownlist, x160 y%y% w475 vdropdownCHDInfo gselectCHDInfo altsubmit, % getDDCHDInfoList()
			y-=5
			gui 3: font, s9 Q5 w400 c000000
			gui 3: add, button, x+15 y%y% w120 h30 gselectCHDInfo vbuttonCHDInfoRight hwndbuttonCHDInfoNextHWND,  % " Next file >"
			imageButton.create(buttonCHDInfoNextHWND, GUI.buttons.default.normal, GUI.buttons.default.hover, GUI.buttons.default.clicked, GUI.buttons.default.disabled) ; Default button colors
			gui 3: font, s12 Q5 w750 c000000
			gui 3: add, text, x285 y350 w200 center border vtextCHDInfoLoading, % "`nLoading...`n"
			gui 3: font, s9 Q5 w400 c000000
			Gui 3:+LastFound +AlwaysOnTop +ToolWindow
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
				
				setTimer, showCHDInfoLoading, -1000 									; Show loading message and disable all elemnts while loading - only if loading is taking longer then normal
				if ( showCHDInfo(job.scannedFiles["info"][CHDInfoFileNum], CHDInfoFileNum, job.scannedFiles["info"].length(), 3) == false ) { 	; Func returned nothing, clear the title
					guiCtrl({textCHDInfoTitle:""})
					guiToggle("hide", "textCHDInfoLoading", 3)
					setTimer, showCHDInfoLoading, off
					return
				}
				setTimer, showCHDInfoLoading, off
				guiToggle("hide", "textCHDInfoLoading", 3) 								; Hide loading message
				guiToggle("enable", ["dropdownCHDInfo", "buttonCHDInfoLeft", "buttonCHDInfoRight"], 3)		; Enable info elements
				guiCtrl("choose", {dropdownCHDInfo:CHDInfoFileNum}, 3)					; Choose dropdown to match newly selected item in CHD info
				controlFocus, ComboBox1, % "CHD Info"									; Keep focus on dropdown to allow arrow keys to also select next and previous files
				
				if ( CHDInfoFileNum == 1 )												; Then disable the appropriate button according to selection number (first or last in list)
					guiToggle("disable", "buttonCHDInfoLeft", 3)						
				else if ( CHDInfoFileNum == job.scannedFiles["info"].length() )
					guiToggle("disable", "buttonCHDInfoRight", 3)
				return
			
			3GUIClose:
				gui 3: destroy
				refreshGUI()
			return

		case "addMeta":
		case "delMeta":
		case default:
	}
	
	if ( !job.workQueue || job.workQueue.length() == 0 ) {																	; Nothing to do!
		log("No jobs found in the work queue")
		msgbox, 16, % "Error", % "No jobs in the work queue!"
		return
	}

	
	job.workTally := {}
	job.availPSlots := []
	job.msgData := []
	job.parseReport := []
	job.allReport := ""
	job.halted := false
	job.started := false
	job.workTally := {started:0, total:job.workQueue.length(), success:0, cancelled:0, skipped:0, withError:0, finished:0, haltedMsg:""}		; Set job variables
	job.workQueueSize := (job.workTally.total < JOB_QUEUE_SIZE)? job.workTally.total : JOB_QUEUE_SIZE											; If number of jobs is less then queue count, only display those progress bars

	toggleMainMenu("hide")																									; Hide main menu bar (selecting menu when running jobs stops messages from being receieved from threads)
	guiToggle("disable", "all")																								; Disable all controls while job is in progress
	guiToggle("hide", "buttonStartJobs")
	guiToggle(["show", "enable"], "buttonCancelAllJobs")
	
	gPos := (job.Cmd == "verify" || job.Cmd == "info")? guiGetCtrlPos("groupboxJob") : guiGetCtrlPos("groupboxOptions")		; Move and show progress bars
	y := gPos.y + gPos.h + 25																								; Assign y (x are from 'gPos') values to groupbox x & y positions
	guiCtrl("moveDraw", {groupBoxProgress:"x5 y" (gPos.y + gPos.h) + 5 " h" job.workQueueSize*25 + 60})						; Move and resize progress groupbox
	guiCtrl("moveDraw", {progressAll:"y" y, progressTextAll: "y" y+4}) 														; Set All Progress bar and it's text Y position
	guiCtrl( {progressAll:0, progressTextAll:"0 jobs of " job.workTally.total " completed - 0%"})
	guiToggle(["show", "enable"], ["groupBoxProgress", "progressAll", "progressTextAll"])									; Show total progress bar
	y += 35
	
	loop % job.workQueueSize {																								; Show individua progress bars																
		guiCtrl("moveDraw", {("progress" a_index):"y" y, ("progressText" a_index):"y" y+4, ("progressCancelButton" a_index):"y" y})	; Move the progress bars into place
		y += 25
		guiCtrl({("progress" a_index):0, ("progressText" a_index): ""})														; Clear the bars text and zero out percentage
		guiToggle(["enable", "show"],["progress" a_index, "progressText" a_index, "progressCancelButton" a_index])			; Enable and show job progress bars
		job.availPSlots.push(a_index)																						; Add available progress slots to queue																						
	}
	
	gui 1:show, autosize																									; Resize main window to fit progress bars
		
	log(job.workTally.total " " stringUpper(job.Cmd) " jobs starting ...")
	SB_SetText(job.workTally.total " " stringUpper(job.Cmd) " jobs started" , 1)

	job.started := true
	job.startTime := a_TickCount
	onMessage(0x004A, "receiveData")											; Receive messages from threads
	
	; Job loop
	while ( job.workTally.finished < job.workTally.total ) { 					; Loop while # of finished jobs are less then job.workTally.total

		if ( job.started == false )
			break
		
		if ( job.availPSlots.length() > 0 && job.workQueue.count() > 0 ) { 		; If there is another queued job and there is a slot available 
			
			thisJob := job.workQueue.removeAt(1)														; Grab the first job from the work queue and assign parameters to variable
			thisJob.pSlot := job.availPSlots.removeAt(1)												; Assign the progress bar a y position from available queue

			job.msgData[thisJob.pSlot] := {}
			
			runCmd := a_ScriptName " threadMode " (SHOW_JOB_CONSOLE == "yes" ? "console" : "")			; "threadmode" flag tells script to run this script as a thread
			run % runCmd ,,, pid																		; Run it
			thisJob.pid := pid

			while ( thisJob.pid <> job.msgData[thisJob.pSlot].pid ) {	; Wait for confirmation that msg was receieved																			
				msg := JSON.Dump(thisJob)
				sendAppMessage(msg, "ahk_class AutoHotkey ahk_pid " pid)
				sleep 250	
			}
			
			job.msgData[thisJob.pSlot].timeout := a_TickCount
		}
		
		loop % job.workQueueSize {			 									; Check for timeouts 
			
			if ( job.msgData[a_index].status == "finished" || job.msgData[a_index].status == "cancelled" )
				continue
			
			if ( (a_TickCount - job.msgData[a_index].timeout) > (TIMEOUT_SEC*1000) ) {					; If timer counter exceeds threshold, we will assume thread is locked up or has errored out 
				
				job.msgData[a_index].status := "error"							; Update job.msgData[] with messages and send "error" flag for that job, then parse the data
				job.msgData[a_index].log := "Error: Job timed out"
				job.msgData[a_index].report := "`nError: Job timed out`n`n`n"
				job.msgData[a_index].progress := 100
				job.msgData[a_index].progressText := "Timed out  -  " job.msgData[a_index].workingTitle
				parseData(job.msgData[a_index])
				
				cancelJob(job.msgData[a_index].pSlot) 							; And attempt to close the process associated with it
			}
		}
		
		sleep 1000
	} ; end job loop


	; Finished all jobs			
	job.started := false
	job.endTime := a_Tickcount
	guiToggle("hide", "buttonCancelAllJobs")
	guiToggle("show", "buttonStartJobs")
	guiToggle("disable", "all")
	
	if ( job.halted ) {																					; There was a fatal error that didnt allow any jobs to be attempted
		log("Fatal Error: " job.workTally.haltedMsg)
		SB_SetText("Fatal Error: " job.workTally.haltedMsg , 1)
		refreshGUI()
		msgBox, 16, % "Fatal Error", % job.workTally.haltedMsg "`n"
	}
	else {																				
		fnMsg := "Total number of jobs attempted: " job.workTally.total "`n"
		fnMsg .= job.workTally.success ? job.FinPreTxt " sucessfully: " job.workTally.success "`n" : ""
		fnMsg .= job.workTally.cancelled ? "Jobs cancelled by the user: " job.workTally.cancelled "`n" : ""
		fnMsg .= job.workTally.skipped ? "Jobs skipped because the output file already exists: " job.workTally.skipped "`n" : ""
		fnMsg .= job.workTally.withError ? "Jobs that finished with errors: " job.workTally.withError "`n" : ""
		fnMsg .= "Total time to finish: " millisecToTime(job.endTime-job.startTime)
		SB_SetText("Jobs finished" (job.workTally.withError ? " with some errors":"!"), 1)
		log( regExReplace(strReplace(fnMsg, "`n", ", "), ", $", "") )
	
		if ( PLAY_SONG_FINISHED == "yes" )																; Play sounds to indicate we are done
			playSound()
		
		msgBox, 36, % APP_MAIN_NAME, % "Finished!`nWould you like to see a report?"
		ifMsgBox Yes
		{
			gui 5: destroy
			gui 5: margin, 10, 20
			gui 5: font, s11 Q5 w700 c000000
			gui 5: add, text,, % job.Desc " report"
			gui 5: font, s9 Q5 w400 c000000
			gui 5: add, edit, readonly y+15 w600 h500, % fnMsg "`n" job.allReport
			gui 5: show, autosize center, REPORT
			Gui 5: +LastFound +AlwaysOnTop +ToolWindow
			controlFocus,, REPORT
			return
		}
		else {
			5guiClose()
			return
		}
	}
}




; User closed the finish window
; ---------------------------------------------------
5guiClose()
{		
	gui 5: destroy
	refreshGUI()
	return
}



; Cancel a single job in progress
; -------------------------------
progressCancelButton()
{
	global job
	gui 1: +ownDialogs
	
	if ( !a_guiControl )
		return
	pSlot := strReplace(a_guiControl, "progressCancelButton", "")
	msgBox, 36,, % "Cancel job " job.msgData[pSlot].idx " - " stringUpper(job.msgData[pSlot].cmd) ": " job.msgData[pSlot].workingTitle "?", 15
	ifMsgBox Yes
		cancelJob(pSlot)
}



; Cancel all jobs currently running
; ---------------------------------
cancelAllJobs()
{
	global job
	
	if ( job.started == false )
		return false
	
	gui 1: +ownDialogs
	msgBox, 36,, % "Are you sure you want to cancel all jobs?", 15
	ifMsgBox No
		return false
	
	loop % job.workQueueSize {	
		cancelJob(a_index)
		sleep 1
	}
	
	loop % job.workQueue.length() {
		thisJob := job.workQueue.removeAt(1)
		job.allReport .= "`n`n" stringUpper(thisJob.cmd) " - " thisJob.workingTitle "`n" drawLine(77) "`n"
		job.allReport .= "Job cancelled by user`n"
		job.workTally.cancelled++
		job.workTally.finished++
		percentAll := ceil((job.workTally.finished/job.workTally.total)*100)
		guiCtrl({progressAll:percentAll, progressTextAll:job.workTally.finished " jobs of " job.workTally.total " completed " (job.workTally.withError ? "(" job.workTally.withError " error" (job.workTally.withError>1? "s)":")") : "")" - " percentAll "%" })
	}
	
	job.workQueue := []										; To make sure we are clear
	job.started := false
	return true	
}



; User cancels job
; --------------------------------
cancelJob(pSlot)
{
	global job
	
	if ( !job.msgData[pSlot].pid || job.msgData[pSlot].status == "finished" )
		return
	
	guiCtrl({("progress" recvData.pSlot):0})
	guiCtrl({("progressText" recvData.pSlot):"Cancelling -  " job.msgData[pSlot].workingTitle})

	job.msgData[pSlot].KILLPROCESS := "true"
	JSONStr := JSON.Dump(job.msgData[pSlot])
	sendAppMessage(JSONStr, "ahk_class AutoHotkey ahk_pid " job.msgData[pSlot].pid)
}



getDDCHDInfoList()
{
	global job
	for idx, filefull in job.scannedFiles["info"]
		ddCHDInfoList .= splitPath(filefull).file "|" ; Create CHD info dropdown list
	return regExReplace(ddCHDInfoList, "\|$")
}


showCHDInfoLoading() 
{
	guiToggle("disable", "all", 3) 
	guiToggle("show", "textCHDInfoLoading", 3) 								; Show loading message
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
		
	file := splitPath(fullFilename)
	guiToggle("enable", "textCHDInfoTitle", guiNum)	
	guiCtrl({"textCHDInfoTitle":"[" (currNum && totalNum ? currNum "/" totalNum "]  " : "") file.file}, guiNum)																	; Change Title to filename
	loop, parse, % runCMD(CHDMAN_FILE_LOC " info -v -i """ fullFilename """", file.dir).msg, % "`n"				; Loop through chdman 'info' stdOut
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
			guiToggle("enable", ["textCHDInfo_" infoLineNum, "editCHDInfo_" infoLineNum], guiNum)
			guiCtrl({("textCHDInfo_" infoLineNum):trim(a[1], " ") ": "}, guiNum)					; Add part 1 as subtitle (ie - "File name", "Size", "SHA1", etc)
			guiCtrl({("editCHDInfo_" infoLineNum):trim(a[2], " ")}, guiNum)								; Part 2 is the information itself 
		}
		else if ( line == "----------  -------  ------------------------------------" )							; When we find this string, we know we are in the overview Compression section
			compressLineNum := 1																				; ... So flag it, use flag as the line counter and move to next loop
		else if ( compressLineNum ) {																					
			line := trim(line, a_space)
			line := regExReplace(line, "      |     |    |   |  ", ";")											; Change all "|" into ";" in line and remove redundant space
			a := strSplit(line, ";")																			; Then split it into part
			if ( a[1] ) {
				guiToggle("enable", ["textCHDInfoHunks_" compressLineNum, "textCHDInfoType_" compressLineNum, "textCHDInfoPercent_" compressLineNum], guiNum)
				guiCtrl({("textCHDInfoHunks_" compressLineNum):trim(a[1], a_space)}, guiNum)								; Part 1 is Hunks
				guiCtrl({("textCHDInfoType_" compressLineNum):trim(a[3] " " a[4] " " a[5] " " a[6], a_space)}, guiNum)	; Part 2 is Compression Type
				guiCtrl({("textCHDInfoPercent_" compressLineNum):trim(a[2], a_space)}, guiNum)							; Part 3 is percentage of compression
			}
			compressLineNum++																							; Add to meta line number
		}
	}
	guiToggle("enable", "editMetadata", guiNum)		; Enable all elements
	guiCtrl({"editMetadata": metadataTxt}, guiNum)
	controlFocus, , % "CHD Info"
	return true
}



; Receieve message data from thread script
; ----------------------------------------
receiveData(data1, data2) 
{
	JSONStr := strGet(numGet(data2 + 2*A_PtrSize) ,, "utf-8")
	data := JSON.Load(JSONStr)
	parseData(data)
}


parseData(recvData) 
{		
	global job, REMOVE_FILE_ENTRY_AFTER_FINISH
	
	job.msgData[recvData.pSlot] := recvData				; Assign globally so we can use anywhere in script - mainly to kill job and check on timeout activity
	job.msgData[recvData.pSlot].timeout := a_Tickcount	; Set to current time -- We know this job is active
	
	if ( recvData.log )
		log("Job " recvData.idx " - " recvData.log)
	
	if ( recvData.report ) {
		if ( !job.parseReport[recvData.idx] )
			job.parseReport[recvData.idx] := "`n`n" stringUpper(recvData.cmd) " - " recvData.workingTitle "`n" drawLine(77) "`n"
		job.parseReport[recvData.idx] .= recvData.report
	}

	switch recvData.status {
		
		case "fileExists":
			job.workTally.skipped++
			SB_SetText("Job " recvData.idx " skipped", 1)
		
		case "error":
			job.workTally.withError++
			SB_SetText("Job " recvData.idx " failed", 1)
			
		case "halted":
			job.halted := true
			job.started := false
			job.workTally.cancelled += job.workQueue.length() + 1		; Tally up totals
			job.workQueue := []											; Empty the work queue
			job.workTally.haltedMsg := recvData.log						; Set flag and error log
			log("Fatal Error. Halted all jobs")
			
		case "success":
			job.workTally.success++
			SB_SetText("Job " recvData.idx " finished successfully!", 1)
			if ( REMOVE_FILE_ENTRY_AFTER_FINISH == "yes" ) {
				removeFromArray(recvData.fromFileFull, job.scannedFiles[recvData.cmd])
				loop % LV_GetCount() {												; Clear finished files from scanned files
					LV_GetText(rtn, a_index)
					if ( rtn == recvData.fromFileFull ) {
						LV_Delete(a_index)
						break
					}
				}
			}
		
		case "cancelled":
			job.workTally.cancelled++
			SB_SetText("Job " job.msgData[pSlot].idx " cancelled", 1)
		
		case "finished":
			job.allReport .= job.parseReport[recvData.idx]
			job.workTally.finished++
			percentAll := ceil((job.workTally.finished/job.workTally.total)*100)
			guiCtrl({progressAll:percentAll, progressTextAll:job.workTally.finished " jobs of " job.workTally.total " completed " (job.workTally.withError ? "(" job.workTally.withError " error" (job.workTally.withError>1? "s)":")") : "")" - " percentAll "%" })
			
			job.availPSlots.push(recvData.pSlot) ; Add an available slot to progress bar array
	}
	
	if ( recvData.progress <> "" )
		guiControl,1:, % "progress" recvData.pSlot, % recvData.progress
	if ( recvData.progressText )
		guiControl,1:, % "progressText" recvData.pSlot, % recvData.progressText
}







; Create  or add to the input files queue (return a work queue)
; -------------------------------------------------------
createJob(command, theseJobOpts, outputExts="", inputExts:="", inputFiles="") 
{
	global
	local wCount :=0, wQueue := [], dupFound := {}, idx, idx2, obj, thisOpt, optVal, cmdOpts := "", fromFileFull, splitFromFile, toExt, q, PID := dllCall("GetCurrentProcessId"), renameDup := 0

	gui 1:submit, nohide
	
	for idx, thisOpt in (isObject(theseJobOpts) ? theseJobOpts : [])								; Parse through supplied Options associated with job
	{
		if ( guiCtrlGet(thisOpt.name "_checkbox", 1) == 0 )											; Skip if the checkbox is not checked
			continue
		if ( thisOpt.editField )
			optVal := guiCtrlGet(thisOpt.name "_edit")
		else if ( thisOpt.dropdownOptions ) {
			optVal := guiCtrlGet(thisOpt.name "_dropdown")											; Get the dropdown value for the current GUI.chdmanOpt
			optVal := isObject(thisOpt.dropdownValues) ? thisOpt.dropdownValues[optVal] : optVal	; If this dropdown option contains a dropdownValues array, optVal becomes the index for that array
		}
		if ( thisOpt.paramString ) {
			optVal := optVal ? (thisOpt.useQuotes ? " """ optVal """" : " " optVal) : ""
			cmdOpts .= " -" thisOpt.paramString . optVal 											; Create the chdman options string
		}
	}
	
	for idx, fromFileFull in (isObject(inputFiles) ? inputFiles : [inputFiles]) {
		splitFromFile := splitPath(fromFileFull)
		
		for idx, toExt in (isObject(outputExts) ? outputExts : [outputExts]) {
			
			q := {}				
			q.idx				:= wQueue.length() + 1
			q.id				:= command q.idx
			q.hostPID			:= PID
			q.cmd				:= command
			q.cmdOpts			:= cmdOpts
			q.inputFileTypes    := isObject(inputExts) ? inputExts : [inputExts]
			q.deleteInputDir	:= deleteInputDir_checkbox
			q.deleteInputFiles 	:= deleteInputFiles_checkbox
			q.keepIncomplete 	:= keepIncomplete_checkbox
			q.workingDir		:= splitFromFile.dir
			q.fromFileExt		:= splitFromFile.ext
			q.fromFile			:= splitFromFile.file
			q.fromFileNoExt 	:= splitFromFile.noExt
			q.fromFileFull		:= fromFileFull
			if ( command <> "verify" && command <> "info" ) {
				q.toFileNoExt	:= outputExts.length()>1 ? splitFromFile.noExt " (" stringUpper(toExt) ")" : splitFromFile.noExt									; For the target file, we use the same base filename as the source
				q.outputFolder	:= createSubDir_checkbox ? OUTPUT_FOLDER "\" q.toFileNoExt : OUTPUT_FOLDER
				q.toFileExt 	:= toExt
				q.toFile		:= q.toFileNoExt "." toExt
				q.toFileFull	:= q.outputFolder "\" q.toFileNoExt "." toExt
				
				; If a duplicate filename was found (ie - 'D:\folder\gameX.chd' and 'C:\folderA\gameX.chd' would both output 'gameX.cue, gameX.bin') ...
				; .. we will suffix a number to the filename gameX-1.chd and gameX-2.chd
				for idx, obj in wQueue {
					if ( obj.toFileFull == q.toFileFull ) {
						dupFound[q.toFileFull] ? dupFound[q.toFileFull]++ : dupFound[q.toFileFull] := 2
						if ( renameDup < 2 ) {
							setTimer, changeMsgBoxButtons, 50
							msgbox, 35, % "Duplicate filename", % "A duplicate conversion was found.`n`nThe file """ q.fromFileFull """  would create a " stringUpper(toExt) " file that has the same name as another file already in the job queue.`n`nSelect [YES] to rename the output file`n""" q.toFile """`nto`n""" obj.toFile " [#" dupFound[q.toFileFull] "]""`n`nSelect [NO] to skip this job"
							ifMsgBox Cancel 	; Rename All 
								renameDup := 2
							ifMsgBox Yes		; Rename
								renameDup := 1
							ifMsgBox No 		; Skip
								renameDup := 0
						}
						if ( renameDup ) {
							q.toFileNoExt	.= " [#" dupFound[q.toFileFull] "]"
							q.outputFolder	:= createSubDir_checkbox ? OUTPUT_FOLDER "\" q.toFileNoExt : OUTPUT_FOLDER
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
			}
			if ( q.count() > 0 ) {
				q.workingTitle 	:= q.toFile ? q.toFile : q.fromFile
				wQueue.push(q) ; Push data to array
			}
		}
	}
	return wQueue
	
	changeMsgBoxButtons:
		if ( !winExist("Duplicate filename") )
			return 
		setTimer, changeMsgBoxButtons, Off 
		winActivate
		controlSetText, Button1, &Rename
		controlSetText, Button2, &Skip
		controlSetText, Button3, Rename &All
	return
}



; Create the main GUI
; -------------------
createMainGUI() 
{
	global
	local idx, key, opt, optName, obj, btn, array := [], ddList := ""
	
	gui 1:-DPIScale 		; hacky workaround to at least get the options crammed in view - for those folks who use higher DPI settings
	
	gui 1:add, button, 		hidden default h0 w0 y0 y0 geditOutputFolder,		; For output edit field (default button
	
	gui 1:add, statusBar
	SB_SetParts(640, 175)
	SB_SetText("  namDHC v" CURRENT_VERSION " for CHDMAN", 2)

	gui 1:add, groupBox, 	x5 w800 h425 vgroupboxJob, Job

	gui 1:add, text, 		x15 y30, % "Job type:"
	
	for key, obj in GUI.dropdowns.job					; Get job dropdown list from object-dictionary
		array[obj.pos] := obj.desc		
	for idx, opt in array								; Order list as order specified by 'pos' in object
		ddList .= opt (a_index == 1  ? "||" : "|")	
	
	gui 1:add, dropDownList,x+5 y28 w200 vdropdownJob gselectJob, % ddList

	gui 1:add, text, 		x+30 y30, % "Media type:"
	gui 1:add, dropDownList,x+5 y28 w200 vdropdownMedia gselectMedia, % "" 	; Media dropdown will be populated when selcting a job

	gui 1:add, text, 		x15 y65, % "Input files"
	
	gui 1:add, button, 		x15 y83 w80 h22 vbuttonAddFiles gaddFolderFiles hwndGUIbutton2, % "Add files"
	gui 1:add, button, 		x+5 y83 w90 h22 vbuttonAddFolder gaddFolderFiles hwndGUIbutton3, % "Add a folder"
	
	gui 1:add, text, 		x455 y93, % "Input file types: "
	gui 1: font, Q5 s9 w700 c000000
	gui 1:add, text, 		x+3 y93 w130 vInputExtTypesText, % ""
	gui 1: font, Q5 s9 w400 c000000
	gui 1:add, button,		x663 y83 w130 h22 gbuttonExtSelect vbuttonInputExtType hwndGUIbutton1, % "Select input file types"
	
	gui 1:add, listView, 	x15 y110 w778 h153 vlistViewInputFiles glistViewInputFiles altsubmit, % "File"
	
	gui 1:add, button, 		x15 y267 w90 vbuttonSelectAllInputFiles gselectInputFiles hwndGUIbutton5, % "Select all"
	gui 1:add, button, 		x+5 y267 w90 vbuttonClearInputFiles gselectInputFiles hwndGUIbutton6, % "Clear selection"
	gui 1:add, button, 		x+20 y267 w90 vbuttonRemoveInputFiles gselectInputFiles hwndGUIbutton4, % "Remove selection"

	gui 1:add, text, 		x15 y305, % "Output Folder"
	gui 1:add, button, 		x15 y324 w90 vbuttonBrowseOutput gcheckNewOutputFolder hwndGUIbutton8, % "Select a folder"
	gui 1:add, text, 		x455 y335,% "Output file types: "
	gui 1: font, Q5 s9 w700 c000000
	gui 1:add, text, 		x+3 y335 w75 vOutputExtTypesText, % ""
	gui 1: font, Q5 s9 w400 c000000
	gui 1:add, button,		x663 y324 w130 h24 vbuttonOutputExtType gbuttonExtSelect hwndGUIbutton7, % "Select output file type"
	
	gui 1:add, edit, 		x15 y352 w778 veditOutputFolder geditOutputFolder, % OUTPUT_FOLDER
	
	gui 1:add, button,		x320 y385 w160 h35 vbuttonStartJobs gbuttonStartJobs hwndstartButtonHWND, % "Start all jobs!"
	
	gui 1:add, button,		hidden x320 y385 w160 h35 vbuttonCancelAllJobs gcancelAllJobs hwndcancelButtonHWND, % "CANCEL ALL JOBS"
	imageButton.create(cancelButtonHWND, GUI.buttons.cancel.normal, GUI.buttons.cancel.hover, GUI.buttons.cancel.clicked)

	gui 1:add, groupBox, 	x5 w800 y435 vgroupboxOptions, % "CHDMAN Options"		; Position and height will be set in refreshGUI()

	loop 8 {	; Stylize default buttons
		btn := "GUIbutton" a_index
		imageButton.create(%btn%, GUI.buttons.default.normal, GUI.buttons.default.hover, GUI.buttons.default.clicked, GUI.buttons.default.disabled) ; Default button colors
	}
	
	for key, opt in GUI.chdmanOpt
	{
		if ( opt.hidden == true )
			continue
		optName := opt.name
		gui 1:add, checkbox,		hidden w200 gcheckboxOption -wrap v%optName%_checkbox,	; Options are moved to their positions when refreshGUI(true) is called
		gui 1:add, edit,			hidden w165 v%optName%_edit,
		gui 1:add, dropdownList, 	hidden w165 altsubmit v%optName%_dropdown,				; ... so we can use for dropdown list to place at same location (default is hidden)
	}

}


; Create GUI progress bar section
; -------------------------------
createProgressBars()
{
	global
	local btn
	
	gui 1:add, groupBox, w800 vgroupBoxProgress, Progress

	gui 1: font, Q5 s9 w700 cFFFFFF
	gui 1:add, progress, hidden x20 w770 h22 backgroundAAAAAA vprogressAll cgreen, 0		; Progress bars y values will be determined with refreshGUI()
	gui 1:add, text,	 hidden x30 w750 h22 +backgroundTrans -wrap vprogressTextAll

	loop % JOB_QUEUE_SIZE_LIMIT {																; Draw but hide all progress bars - we will only show what is called for later
		gui 1:add, progress, hidden x20 w740 h22 backgroundAAAAAA vprogress%a_index% c17A2B8, 0				
		gui 1:add, text,	 hidden x30 w720 h22 +backgroundTrans -wrap vprogressText%a_index%
		gui 1:add, button,	 hidden x+15 w25 vprogressCancelButton%a_index% gprogressCancelButton hwndprogCancelbutton%a_index%, % "X"
		btn := "progCancelbutton" a_index
		imageButton.create(%btn%, GUI.buttons.cancel.normal,	GUI.buttons.cancel.hover, GUI.buttons.cancel.clicked)
	}
}
 

; Create GUI Menus
; ----------------
createMenus() 
{
	global GUI, JOB_QUEUE_SIZE_LIMIT, JOB_QUEUE_SIZE
	
	loop % JOB_QUEUE_SIZE_LIMIT
		menu, SubSettingsConcurrently, Add, %a_index%, % "menuSelected"
	menu, SubSettingsConcurrently, Check, % JOB_QUEUE_SIZE						; Select current jobQueue number

	loop % GUI.menu.namesOrder.length() {
		menuName := GUI.menu.namesOrder[a_index]
		menuArray := GUI.menu[menuName]
		
		loop % menuArray.length() {
			menuItem :=  menuArray[a_index]
			menu, % menuName "Menu", add, % menuItem.name, % menuItem.gotolabel
		
			if ( menuItem.saveVar ) {
				saveVar := menuItem.saveVar
				menu, % menuName "Menu", % (%saveVar% == "yes"? "Check":"UnCheck"), % menuItem.name
			}
		}
		menu, MainMenu, add, % menuName, % ":" menuName "Menu"
	}
	gui 1:menu, MainMenu
	
	menu, % "InputExtTypes", add				; Input & Output extension dummy menus to populate with refreshGUI() later				
	menu, % "OutputExtTypes", add
	
	Menu, Tray, NoStandard						; Remove default options in tray icon
	Menu, Tray, Add, E&xit, quitApp
}



; Show or hide main menu
; -------------------------------
toggleMainMenu(showOrHide:="show")
{
	global mainAppHWND
	static visble, hMenu
	
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
	static winCreated
	
	if ( !winCreated ) {
		winCreated := true
		gui 2:-sysmenu +resize
		gui 2:margin, 5, 10
		gui 2:add, edit, % "w" APP_VERBOSE_WIN_WIDTH-10 " h" APP_VERBOSE_WIN_HEIGHT-20 " readonly veditVerbose",
	}
	if ( show == "yes" ) {
		gui 2:show, % "w" APP_VERBOSE_WIN_WIDTH " h" APP_VERBOSE_WIN_HEIGHT " x" APP_VERBOSE_WIN_POS_X " y" APP_VERBOSE_WIN_POS_Y, % APP_VERBOSE_NAME
		sendMessage 0x115, 7, 0, Edit1, % APP_VERBOSE_NAME		; Scroll to bottom of log
		controlFocus,, % APP_VERBOSE_NAME							; Removes seletced text effect most times
		controlClick, Edit1, % APP_VERBOSE_NAME 					; Removes seletced text effect when showing window first time
	}
	else if ( show == "no" )
		gui 2:hide
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
	sendMessage 0x115, 7, 0, Edit1, % APP_VERBOSE_NAME	; Scroll to bottom of log
}



; Read or write to ini file
; -------------------------
ini(job="read", var:="") 
{
	global
	local varsArry := isObject(var)? var : [var], idx, varName
	
	if ( varsArry[1] == "" )
		return false

	for idx, varName in varsArry {
		if ( job == "read" ) {
			defaultVar := %varName%
			iniRead, %varName%, % APP_MAIN_NAME ".ini", Settings, % varName
			if ( %varName% == "ERROR" || %varName% == "" ) {
				%varName% := defaultVar
			}
		}
		else if ( job == "write" ) {
			if ( %varName% == "ERROR" || %varName% == "" )
				%varName% := %varName%
			iniWrite, % %varName%, % APP_MAIN_NAME ".ini", Settings, % varName
			;log("Saved " varName " with value " %varName%)
		}
	}
}



playSound() 
{
	if ( inStr(A_OSVersion, "10") || inStr(A_OSVersion, "11") )
		play := [a_WinDir "\Media\Alarm05.wav", 1]
	else if ( inStr(A_OSVerson, "7") )
		play := [a_WinDir "\Media\notify.wav", 2, "wait"]
	
	if ( play ) {
		loop % play[2] {
			soundPlay, % play[1], % play[3]
			if ( play[2] > 1 )
				sleep 100
		}
	}
	else {
		SoundBeep, 300, 200
		SoundBeep, 600, 800
		sleep 500
		SoundBeep, 300, 200
		SoundBeep, 600, 800
	}
}



; Send data across script instances
; -------------------------------------------------------
sendAppMessage(stringToSend, targetScriptTitle) 
{
  VarSetCapacity(CopyDataStruct, 3*A_PtrSize, 0)
  SizeInBytes := strPutVar(stringToSend, stringToSend, "utf-8")
  NumPut(SizeInBytes, CopyDataStruct, A_PtrSize)
  NumPut(&stringToSend, CopyDataStruct, 2*A_PtrSize)
  Prev_DetectHiddenWindows := A_DetectHiddenWindows
  Prev_TitleMatchMode := A_TitleMatchMode
  DetectHiddenWindows On
  SetTitleMatchMode 2
  SendMessage, 0x4a, 0, &CopyDataStruct,, % targetScriptTitle
  DetectHiddenWindows %Prev_DetectHiddenWindows%
  SetTitleMatchMode %Prev_TitleMatchMode%
  return errorLevel
}

strPutVar(string, ByRef var, encoding)
{
	varSetCapacity( var, StrPut(string, encoding) * ((encoding="utf-8"||encoding="cp1200") ? 2 : 1) )
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
	if ( fileExist(newFolder) <> "D" ) {						; Folder dosent exist
		if ( !splitPath(newFolder).drv ) {						; No drive letter can be assertained, so it's invalid
			newFolder := false
		} else {												; Output folder is valid but dosent exist
			fileCreateDir, % regExReplace(newFolder, "\\$")
			newFolder := errorLevel ? false : newFolder
		}
	}
	return newFolder											; Returns the folder name if created or it exists, or false if no folder was created
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
showObj(obj, s := "", show:=true) 
{
	static str
	if (s == "")
		str := ""
	for idx, v In obj {
		n := (s == "" ? idx : s . ", " . idx)
		if isObject(v)
			showObj(v, n, false)
		else
			str .= "[" . n . "] = " . v . "`r`n"
	}
	
	if ( show ) {
		gui showObj: +resize +toolwindow
		gui showObj: margin, 10, 10
		gui showObj: add, edit, readonly w800 h600, % str
		gui showObj: show, autosize, % "SHOW OBJECT"
		WinWaitClose, % "SHOW OBJECT"
	}
	
	return rTrim(str, "`r`n")
}


procCountDDList()
{
	loop % envGet("NUMBER_OF_PROCESSORS")
		lst .= a_index "|"					; Create processor count dropdown list
	return "|" lst "|"						; Last "|" is to select last as default
}

; Splitpath function
; -------------------
splitPath(inputFile) 
{
	splitPath inputFile, file, dir, ext, noext, drv
	return {full:inputFile, file:file, dir:dir, ext:ext, noext:noext, drv:drv}
}


; To allow an inline call of the LV_GetText() function
; -----------------------------------------------------
LV_GetText2(row, byRef rtn="")				
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
	stringUpper, rtn, str, % title ? "T":""
	return rtn
}


; Join an array wilth delimiters and return a string
strJoin(arr, delim) { 
    result := ""
    for each, val in arr
        result .= val . delim
    return rTrim(result, delim)
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
arrayToString(thisArray, delim:=", ")
{
	for idx, val in thisArray
		rtn .= val delim
	return regExReplace(rtn, delim, "", "", 1, -1)
}


; Create an Array from a string with delimeters
; ---------------------------------------------
arrayFromString(string, delim:=",") 
{
	rtnArray := []
	loop parse, string, % delim
		rtnArray.push(a_loopfield)
	return rtnArray
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
; by SKAN
; ------------------------
guiDefaultFont() 
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
	global APP_MAIN_WIN_POS_X, APP_MAIN_WIN_POS_Y, APP_VERBOSE_WIN_POS_X, APP_VERBOSE_WIN_POS_Y, APP_MAIN_NAME, APP_VERBOSE_NAME
	
	if ( a_gui == 1 ) {
		winGetPos, APP_MAIN_WIN_POS_X, APP_MAIN_WIN_POS_Y,,, % APP_MAIN_NAME
		setTimer, writemoveGUIWin, -500
	}
	else if ( a_gui == 2 ) {
		winGetPos, APP_VERBOSE_WIN_POS_X, APP_VERBOSE_WIN_POS_Y,,, % APP_VERBOSE_NAME
		setTimer, writemoveGUIWin, -500
	}
	return
	
	writemoveGUIWin:
		ini("write", ["APP_MAIN_WIN_POS_X", "APP_MAIN_WIN_POS_Y", "APP_VERBOSE_WIN_POS_X", "APP_VERBOSE_WIN_POS_Y"])
	return
}


; Verbose window was resized
; --------------------------
2GuiSize(guiHwnd, eventInfo, W, H) 
{
	global APP_VERBOSE_WIN_HEIGHT := H, APP_VERBOSE_WIN_WIDTH := W
	
	autoXYWH("wh", "editVerbose") 						; Resize edit control with window
	setTimer, write2GuiSize, -500					
	return
	
	write2GuiSize:
		ini("write", ["APP_VERBOSE_WIN_HEIGHT", "APP_VERBOSE_WIN_WIDTH"])
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
} } 
}


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


; guiToggle GUI controls
; ------------------------------------------
guiToggle(doWhat, whichControls, guiNum:=1) 
{
	global APP_MAIN_NAME
	
	if ( !doWhat || !whichControls )
		return false
	
	doWhatArray := isObject(doWhat) ? doWhat : [doWhat]
	ctlArray := isObject(whichControls) ? whichControls : (whichControls == "all" ? getWinControls(APP_MAIN_NAME, "Static") : [whichControls])

	for idx, dw in doWhatArray
		for idx2, ctl in ctlArray
			guiControl %guiNum%:%dw%, % ctl
}



; Function replacement for guiControlGet
; --------------------------------------
guiCtrlGet(ctrl, guiNum:=1) 
{
	guiControlGet, rtn, %guiNum%:, %ctrl%
	return rtn
}



; Get a windows control elements as an array
; -------------------------------------------
getWinControls(win, ignoreStr:="") 
{
	rtnArray := []
	winGet, ctrList, ControlList, % win
	loop, parse, ctrList, `n
	{
		if ( ignoreStr && inStr(a_loopfield, ignoreStr) ) ; Dont disable text elements
			continue
		rtnArray.push(a_loopfield)	
	}
	return rtnArray
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
		rtn .= "-"
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
fileDelete(file, attempts:=5, sleepdelay:=50) 
{
	f := fileExist(file)
	if ( !f || f == "D" ) 									; if file dosent exist or file is a directory 
		return true
	
	loop % (attempts < 1 ? 1 : attempts) { 					; 5 attempts to delete the file
		fileDelete, % file
		sleep % sleepdelay
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
	if ( fileExist(dir) <> "D" )								; If supplied dir isn't a directory, we are good to go
		return true
	;if ( dllCall("Shlwapi\PathIsDirectoryEmpty", "Str", dir) ) ; if empty
	loop % (attempts < 1 ? 1 : attempts) {
		fileRemoveDir % dir, % full								; Attempt to delete the directory x times, full flag for when folder is full
		sleep % sleepdelay	
		if ( fileExist(dir) <> "D" )							; Success
			return true
	}
	return false
}


deleteFilesReturnList(file) 
{
	delFiles := ""
	for idx, thisFile in getFilesFromCUEGDITOC(file)
		delFiles .= (fileDelete(thisFile, 3, 100) ? thisFile ", " : "")
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
		f := splitPath(thisFile)
		loop, Read, % thisFile 
		{
			switch f.ext {
				case "cue", "toc":
					if ( stPos := inStr(a_loopReadLine, "FILE """, true) ) {
						stPos += 6
						endPos := inStr(a_loopReadLine, """", true, -1)
						file := subStr(a_loopReadLine, stPos, (endPos-stPos))
						if ( fileExist(f.dir "\" file) )
							fileList.push(f.dir "\" file)
					}
				case "gdi":
					if ( a_loopReadLine is digit && a_index > 1 ) {
						loop parse, a_loopReadLine, " ;"
							if ( fileExist(f.dir "\" a_loopField) )
								fileList.push(f.dir "\" a_loopField)
					}
			}
		}
	}
	return fileList
}


; Unzip a file
;http://www.autohotkey.com/forum/viewtopic.php?p=402574
; -----------------------------------------------------
unzip(sZip, sUnz)
{
    try {
		fso := ComObjCreate("Scripting.FileSystemObject")
		If Not fso.FolderExists(sUnz)  
		   fso.CreateFolder(sUnz)
		psh  := ComObjCreate("Shell.Application")
		zippedItems := psh.Namespace( sZip ).items().count
		psh.Namespace( sUnz ).CopyHere( psh.Namespace( sZip ).items, 4|16 )
		Loop {
			sleep 50
			unzippedItems := psh.Namespace( sUnz ).items().count
			IfEqual, zippedItems, %unzippedItems%
				break
		}
		return true
	}
	catch e {
		msgbox, 16, % "Unzipping Error", % "Unzipping Error:`n" e
		return false
	}
}


; Read a zip file
; ----------------
readZipFile(zipFile) 
{
	if ( splitPath(zipFile).ext <> "zip" )
		return false
	array := []
	zippedItem := comObjCreate("Shell.Application").Namespace(zipFile)
	for file, val in zippedItem.items
		array.push(file.Name)
	return array
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


; Thanks maestrith 
; https://www.autohotkey.com/board/topic/88685-download-a-url-to-a-variable/
URLDownloadToVar(url){
	try {
		hObject:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
		hObject.Open("GET",url)
		hObject.Send()
		return hObject.ResponseText
	}
}


normalizePath(path) {
    cc := DllCall("GetFullPathName", "str", path, "uint", 0, "ptr", 0, "ptr", 0, "uint")
    VarSetCapacity(buf, cc*2)
    DllCall("GetFullPathName", "str", path, "uint", cc, "str", buf, "ptr", 0)
    return buf
}


; Check github for for newest assets
; ----------------------------------
checkForUpdates(arg1:="", userClick:=false) 
{
	global CURRENT_VERSION, APP_MAIN_NAME, GITHUB_REPO_URL, DIR_TEMP
	gui 4:+OwnDialogs
	
	log("Checking for updates ... ")
	
	if ( !a_isCompiled ) {
	 	if ( userClick )
			msgbox 16, % "Error", % "Can only update compiled binaries"
		log("Error updating: Can only update compiled binaries")  ; no time-stamp
		return
	}
		
	/*
	obj.tag_name 						= version (ie-"namDHCv1.03")
	obj.body							= version changes
	obj.assets[1].browser_download_url	= URL *should* point to chdman.exe
	obj.assets[2].browser_download_url	= URL *should* point to namDHC.exe
	obj.assets[3].browser_download_url	= URL *should* point to namDHC_vx.xx.zip
	obj.created_at						= date created
	*/	
	JSONStr := URLDownloadToVar(GITHUB_REPO_URL)
	obj := JSON.Load(JSONStr)
	
	if ( !isObject(obj) ) {
		log("Error updating: Update info invalid")
		if ( userClick )
			msgbox 16, % "Error getting update info", % "Update info invalid"
	}
	else if ( obj.message && inStr(obj.message, "limit exceeded") ) {
		log("Error updating: Github API limit exceeded")
		if ( userClick )
			msgbox 16, % "Error getting update info", % "Github API limit exceeded"
	}
	else if ( !obj.tag_name ) {
		log("Error updating: Update info invalid")
		if ( userClick )
			msgbox 16, % "Error getting update info", % "Update info invalid"
	}	
	else {
		newVersion := regExReplace(obj.tag_name, "namDHC|v| ", "")

		if ( newVersion == CURRENT_VERSION ) {
			log("No new updates found. You are running the latest version")
			if ( userClick )
				msgbox 64, % "No new updates found", % "You are running the latest version"
			return
		}
		
		else if ( newVersion < CURRENT_VERSION ) {
			log("Your version is newer then the latest release!  Current version: v" CURRENT_VERSION " - Latest version: v" newVersion)
			if ( userClick )
				msgbox 16, % "Error", % "Your version is newer then the latest release!`n`nCurrent version: v" CURRENT_VERSION " - Latest version: v" newVersion
		}
		
		else if ( newVersion > CURRENT_VERSION ) {
			log("An update was found: v" newVersion)
			msgBox, 36, % "Update available", % "A new version of " APP_MAIN_NAME " is available!`n`nCurrent version: v" CURRENT_VERSION "`nLatest version: v" newVersion "`n`nChanges:`n" strReplace(obj.body, "-", "    -") "`n`nDo you want to update?"
			ifMsgBox No
				return
			
			for idx, asset in obj.assets {
				if ( inStr(asset.browser_download_url, "namDHC.exe") )
					namDHCBinURL := asset.browser_download_url
				if ( inStr(asset.browser_download_url, "chdman.exe") )
					chdmanBinURL := asset.browser_download_url
				if ( inStr(asset.browser_download_url, ".zip") )
					chdmanZipURL := asset.browser_download_url	
			}
			if ( !namDHCBinURL ) {
				msgbox 16, % "Error", % "Update not found!"
				log("Error updating: Update binary couldn't be found in github repo!")
				return
			}
			
			fileTemp := DIR_TEMP "\namDHC.exe"
			batchFile := DIR_TEMP "\update.bat"
			batchText := "@timeout /t 1 /nobreak > NUL`r`n@del """ a_ScriptFullPath """ > NUL`r`n@copy """ fileTemp """ """ a_ScriptFullPath """ > NUL`r`n@start " a_ScriptFullPath "`r`n@exit 0`r`n"
			
			createFolder(DIR_TEMP)
			fileDelete(fileTemp, 3, 100) 						; delete if temp file already exists
			
			urlDownloadToFile, % namDHCBinURL, % fileTemp
			if ( !fileExist(fileTemp) ) {
				msgbox 16, % "Error", % "Error downloading update!"
				log("Error updating: There was an error downloading the update")
				return
			}
			
			fileDelete(batchFile, 3, 100) ; Delete before creating a new batch file
			fileAppend, % batchText, % batchFile
			sleep 50
			
			run % batchFile
			exitApp
		}
	}
}




; Kill all namDHC process (including chdman.exe)
; -----------------------------------------------
killAllProcess() 
{
	global APP_MAIN_NAME, APP_VERBOSE_NAME, APP_RUN_JOB_NAME, APP_RUN_CONSOLE_NAME

	loop {
		process, close, % "chdman.exe"
		if ( errorLevel == 0 )
			break
	}

	for idx, app in [APP_MAIN_NAME, APP_VERBOSE_NAME, APP_RUN_JOB_NAME, APP_RUN_CONSOLE_NAME] {
		hwnd := winExist(app)
		winActivate % "ahk_id " hwnd
		winClose % "ahk_id " hwnd
		if ( winExist("ahk_id " hwnd) ) {
			postMessage, 0x0112, 0xF060,,, % "ahk_id " hwnd
			winKill % "ahk_id " hwnd
		}
	}
}


; Close App
; ---------
GuiClose()
{
	global
	
	if ( job.started == true ) {
		if ( cancelAllJobs() == false )
			return 1
		else {
			refreshGUI()
			guiToggle("disable", "all")
		}
	}
	exitApp
}


quitApp() 
{
}
#Warn VarUnset, Off
OnMessage(0x4A, (wParam, lParam, msg, hwnd) => thread_receiveData(wParam, lParam))

threadConsole := 0
gui1 := Gui('-SysMenu', APP_RUN_JOB_NAME)
;menu tray, noIcon

APP_MAIN_HWND := winExist(APP_MAIN_NAME)
threadRecvData := {cmd:"", idx:""}
threadJobData := {}
threadSendObj := {}
threadJobWatch := {active:false, cmd:"", outputFile:"", startSize:0, lastPct:-1, lastTick:0}

thread_log("Started ... ")

; Set up console window
; --------------------------------------------------------------------------------------------
threadConsole := Console()
threadConsole.setConsoleTitle(APP_RUN_CONSOLE_NAME)
threadConsoleHWND := threadConsole.getConsoleHWND()

if ( threadConsoleHWND )
	WinSetStyle("^0x80000", "ahk_id " threadConsoleHWND) 					; Remove close button on window

if ( SHOW_JOB_CONSOLE == "no" ) 												; Hide console
	if ( threadConsoleHWND )
		WinHide("ahk_id " threadConsoleHWND)

; Handshaking
; --------------------------------------------------------------------------------------------
thread_log("Handshaking with namDHC... ")

while ( !threadRecvData.cmd || !threadRecvData.idx ) {
	thread_log(".")
	Sleep(10)
	if ( a_index > 1000 ) {
		thread_log("Error handshaking!`n")
		return
	}
}
recvCmd := threadRecvData.HasOwnProp("cmd") ? threadRecvData.cmd : ""
recvWorkingTitle := threadRecvData.HasOwnProp("workingTitle") ? threadRecvData.workingTitle : "Unknown job"
recvWorkingDir := threadRecvData.HasOwnProp("workingDir") ? threadRecvData.workingDir : A_ScriptDir
thread_log("OK!`n"
		. "Starting a " stringUpper(recvCmd) " job`n`n"
		. "Working job file: " recvWorkingTitle "`n"
		. "Working directory: " recvWorkingDir "`n")

mergeObj(threadRecvData, threadJobData)
mergeObj(threadRecvData, threadSendObj)															; Assign threadRecvData to threadSendData as we will be sending the same info back and forth

threadSendObj.pid := dllCall("GetCurrentProcessId")
threadSendObj.progress := 0
threadSendObj.log := "Preparing " stringUpper(recvCmd) " - " recvWorkingTitle
threadSendObj.progressText := "Preparing  -  " recvWorkingTitle
threadSendData()


; Create output folder if it dosent exist
; -------------------------------------------------------------------------------------------------------------------
recvOutputFolder := threadRecvData.HasOwnProp("outputFolder") ? threadRecvData.outputFolder : ""
recvFromFileExt := threadRecvData.HasOwnProp("fromFileExt") ? StringLower(threadRecvData.fromFileExt) : ""
recvFromFile := threadRecvData.HasOwnProp("fromFile") ? threadRecvData.fromFile : ""
recvFromFileNoExt := threadRecvData.HasOwnProp("fromFileNoExt") ? threadRecvData.fromFileNoExt : ""
recvFromFileFull := threadRecvData.HasOwnProp("fromFileFull") ? threadRecvData.fromFileFull : ""
recvToFileFull := threadRecvData.HasOwnProp("toFileFull") ? threadRecvData.toFileFull : ""
recvCmdOpts := threadRecvData.HasOwnProp("cmdOpts") ? threadRecvData.cmdOpts : ""
recvKeepIncomplete := threadRecvData.HasOwnProp("keepIncomplete") && threadRecvData.keepIncomplete
recvDeleteInputFiles := threadRecvData.HasOwnProp("deleteInputFiles") && threadRecvData.deleteInputFiles
recvDeleteInputDir := threadRecvData.HasOwnProp("deleteInputDir") && threadRecvData.deleteInputDir
threadRecvData.outputFolderCreatedByJob := false
threadJobData.outputFolderCreatedByJob := false

if ( recvOutputFolder ) {
	if ( fileExist(recvOutputFolder) != "D" ) {
		if ( createFolder(recvOutputFolder) ) {
			threadRecvData.outputFolderCreatedByJob := true
			threadJobData.outputFolderCreatedByJob := true
			thread_log("Created directory " recvOutputFolder "`n")
		}
		else {
			Sleep(50)
			thread_log("Error creating directory " recvOutputFolder "`n")
			
			threadSendObj.status := "error"
			threadSendObj.log := "Error creating directory " recvOutputFolder
			threadSendObj.report := "`n" "Error creating directory " recvOutputFolder "`n"
			threadSendObj.progressText := "Error creating directory  -  " recvWorkingTitle
			threadSendObj.progress := 100
			threadSendData()
			thread_finishJob()
			ExitApp()
		}
	}
}

; Archive file was supplied as source
; -------------------------------------------------------------------------------------------------------------------
tempZipDirectory := ""
zipErr := ""
archiveAction := recvFromFileExt == "zip" ? "Unzipping" : "Extracting archive"
archiveDoneAction := recvFromFileExt == "zip" ? "Unzipped" : "Extracted archive"
if ( recvFromFileExt == "zip" || recvFromFileExt == "7z" ) {

	threadSendObj.status := "unzipping"
	threadSendObj.progress := 0
	threadSendObj.progressText := archiveAction "  -  " recvFromFile
	threadSendData()
	
	tempZipDirectory := DIR_TEMP "\" recvFromFileNoExt
	folderDelete(tempZipDirectory, 3, 25, 1) 										; Delete folder and its contents if it exists
	createFolder(tempZipDirectory)													; Create the folder
	
	if ( fileExist(tempZipDirectory) == "D" ) {
		thread_log(archiveAction " " recvFromFileFull "`nExtracting to: " tempZipDirectory)
		SetTimer(thread_timeout, (TIMEOUT_SEC*1000)*3) 							; Set timeout timer timeout time x2
		if ( fileUnzip := thread_extractArchive(recvFromFileFull, tempZipDirectory) ) {
			threadSendObj.log := archiveDoneAction " " recvFromFileFull " successfully"
			threadSendObj.report := archiveDoneAction " " recvWorkingTitle " successfully`n"
			threadSendObj.progress := 100
			threadSendObj.progressText := archiveDoneAction " successfully  -  " recvWorkingTitle
			threadSendData()
		
			thread_log(archiveDoneAction " " recvFromFileFull " successfully`n")		
			
			threadRecvData.unzipped := {}
			threadRecvData.unzipped.fromFileFull := fileUnzip.full
			threadRecvData.unzipped.fromFile := fileUnzip.file
			threadRecvData.unzipped.fromFileNoExt := fileUnzip.noExt
			threadRecvData.unzipped.fromFileExt := fileUnzip.ext
			
			mergeObj(threadRecvData, threadSendObj)
		}
		else zipErr := ["Error extracting archive file '" recvFromFileFull "'", "Error extracting archive  -  " recvFromFile]
	}
	else zipErr := ["Error creating temporary directory '" DIR_TEMP "\" recvFromFileNoExt "'", "Error creating temp directory"]
	
	SetTimer(thread_timeout, 0)
	
	if ( zipErr ) {
		threadSendObj.status := "error"
		threadSendObj.log := zipErr[1]
		threadSendObj.report := "`n" zipErr[1] "`n"
		threadSendObj.progressText := zipErr[2] "  -  " recvWorkingTitle
		threadSendObj.progress := 100
		threadSendData()

		if ( fileExist(tempZipDirectory) )
			thread_deleteDir(tempZipDirectory, 1) ; Delete temp directory
		thread_cleanupCreatedOutputFolder()
		
		thread_finishJob()
		ExitApp()
	}
}

Sleep(10)
	
unzippedFromFileFull := ""
try unzippedFromFileFull := threadRecvData.unzipped.fromFileFull

fromFile := unzippedFromFileFull ? "`"" . unzippedFromFileFull . "`"" : (recvFromFileFull ? "`"" . recvFromFileFull . "`"" : "" )
toFile := recvToFileFull ? "`"" recvToFileFull "`"" : ""
chdmanExe := "`"" CHDMAN_FILE_LOC "`""
cmdLine := chdmanExe . " " . recvCmd . recvCmdOpts . (fromFile ? " -i " fromFile : "") . (toFile ? " -o " toFile : "")
thread_log("`nCommand line: " cmdLine "`n`n")

forceOverwrite := RegExMatch(" " recvCmdOpts " ", "i)\s-f(\s|$)")
if ( toFile ) {
	outputFile := StrReplace(toFile, "`"", "")
	if ( !forceOverwrite && FileExist(outputFile) ) {
		threadSendObj.status := "fileExists"
		threadSendObj.log := "file already exists - skipping"
		threadSendObj.report := "Warning: file already exists - skipping:`n`n" outputFile "`nUse Force (-f) to overwrite.`n"
		if ( recvDeleteInputFiles )
			threadSendObj.report .= "Input files were not deleted because the job was skipped.`n"
		threadSendObj.progressText := "File already exists - Skipping  -  " recvWorkingTitle
		threadSendObj.progress := 100
		threadSendData()
		thread_cleanupCreatedOutputFolder()
		thread_finishJob()
		ExitApp()
	}
}

threadSendObj.progress := 0
threadSendObj.log := "Starting " stringUpper(recvCmd) " - " recvWorkingTitle
threadSendObj.progressText := "Starting job  -  " recvWorkingTitle

threadSendData()


; Get starting file size
fileStartSize := 0
inputFileForSize := strReplace(fromFile, "`"", "")
inputExtForSize := ""
if ( threadRecvData.HasOwnProp("fromFileExt") && threadRecvData.fromFileExt )
	inputExtForSize := StringLower(threadRecvData.fromFileExt)
else if ( inputFileForSize )
	inputExtForSize := StringLower(splitFilePath(inputFileForSize).ext)

if ( inputExtForSize == "gdi" || inputExtForSize == "cue" || inputExtForSize == "toc" ) {
	for _, thisFile in getFilesFromCUEGDITOC(inputFileForSize) { 				; Sum source descriptor + track files
		if ( FileExist(thisFile) )
			fileStartSize += FileGetSize(thisFile)
	}
}
else if ( inputFileForSize && FileExist(inputFileForSize) )
	fileStartSize := FileGetSize(inputFileForSize)  							; Use single-file source size for other media types


timeoutMs := TIMEOUT_SEC*1000
if ( InStr(recvCmd, "create") )
	timeoutMs := timeoutMs < 300000 ? 300000 : timeoutMs			; Create jobs can stay quiet for a while before first progress burst.
SetTimer(thread_timeout, timeoutMs) 								; Set timeout timer
threadJobWatch.active := true
threadJobWatch.cmd := recvCmd
threadJobWatch.outputFile := strReplace(toFile, "`"", "")
threadJobWatch.startSize := fileStartSize
threadJobWatch.lastPct := -1
threadJobWatch.lastTick := A_TickCount
SetTimer(thread_heartbeat, 1000)
output := runCMD(cmdLine, recvWorkingDir, "CP0", thread_parseCHDMANOutput) ; thread_parseCHDMANOutput is the function that will be called for STDOUT 
SetTimer(thread_heartbeat, 0)
threadJobWatch.active := false
SetTimer(thread_timeout, 0)

logicalInputSize := 0
if ( output.HasOwnProp("msg") && output.msg ) {
	if ( RegExMatch(output.msg, "i)Logical size:\h*([\d,]+)", &mLogical) )
		logicalInputSize := Integer(StrReplace(mLogical[1], ",", ""))
	else if ( RegExMatch(output.msg, "i)Input file size:\h*([\d,]+)", &mInput) )
		logicalInputSize := Integer(StrReplace(mInput[1], ",", ""))
}

if ( output.exitcode == -1 && output.HasOwnProp("msg") && output.msg ) {
	; If CHDMAN banner/progress was captured, process started successfully and this is a fallback transport error.
	if ( InStr(output.msg, "Compressed Hunks of Data")
		|| InStr(output.msg, "Output CHD:")
		|| InStr(output.msg, "Compressing")
		|| InStr(output.msg, "Compression complete") )
		output.exitcode := 0
}

if ( output.exitcode == -1 ) {
	threadSendObj.status := "error"
	threadSendObj.log := "Error: Failed to launch CHDMAN process"
	threadSendObj.report := "Error: Failed to launch CHDMAN process.`n`nCommand: " cmdLine
	if ( output.HasOwnProp("msg") && output.msg )
		threadSendObj.report .= "`nDetails: " output.msg
	threadSendObj.report .= "`n"
	threadSendObj.progressText := "Error launching CHDMAN  -  " recvWorkingTitle
	threadSendObj.progress := 100
	threadSendData()
	thread_cleanupCreatedOutputFolder()
	thread_finishJob()
	ExitApp()
}

rtnError := thread_checkForErrors(output.msg)
if ( !rtnError && output.exitcode != 0 )
	rtnError := "CHDMAN exited with code " output.exitcode

; CHDMAN was not successful - Errors were detected
; -------------------------------------------------------------------------------------------------------------------
if ( rtnError ) {
	
	if ( inStr(rtnError, "file already exists") == 0
		&& !recvKeepIncomplete
		&& recvToFileFull ) {			; Delete incomplete output files, only delete files that arent "file exists" error
		
		delFiles := deleteFilesReturnList(recvToFileFull)
		threadSendObj.log := delFiles != "" ? "Deleted incomplete file(s): " regExReplace(delFiles, " ,$") : "Error deleting incomplete file(s)!"
		threadSendObj.report := delFiles != "" ? thread_formatReportList("Deleted incomplete file(s):", delFiles) : threadSendObj.log "`n"
		threadSendObj.progress := 100
		threadSendData()
		
		thread_log(threadSendObj.log "`n")
	}
	
	threadSendObj.status := "error"
	threadSendObj.log := rtnError
	threadSendObj.report := (inStr(threadSendObj.log, "Error") ? "" : "Error: ") threadSendObj.log "`n`n"
	if ( output.HasOwnProp("exitcode") )
		threadSendObj.report .= "Exit code: " output.exitcode "`n"
	if ( output.HasOwnProp("msg") && output.msg )
		threadSendObj.report .= "CHDMAN output:`n" output.msg "`n"
	threadSendObj.progressText := regExReplace(threadSendObj.log, "`n|`r", "") "  -  " recvWorkingTitle
	threadSendObj.progress := 100
	threadSendData()
	thread_cleanupCreatedOutputFolder()
	
	thread_finishJob()
	ExitApp()

}

; CHDMAN was successfull - No errors were detected
; -------------------------------------------------------------------------------------------------------------------
if ( fileExist(tempZipDirectory) ) 
	thread_deleteDir(tempZipDirectory, 1) 														; Always delete temp zip directory and all of its contents

if ( recvDeleteInputFiles ) {
	
	if ( unzippedFromFileFull && fileExist(unzippedFromFileFull) ) 	; Delete input files of unzipped if requested (and they exist)
		thread_deleteFiles(unzippedFromFileFull)

	if ( recvFromFileFull && fileExist(recvFromFileFull) ) 				; Delete input source files if requested
		thread_deleteFiles(recvFromFileFull)
}

if ( recvDeleteInputDir && fileExist(recvWorkingDir) == "D" )			; Delete input folder only if requested and it is not empty
	thread_deleteDir(recvWorkingDir)	

; Include useful verification details in the final report (ignore banner/progress spam).
if ( InStr(recvCmd, "verify")
	&& output.HasOwnProp("msg")
	&& output.msg ) {
	verifyReport := ""
	loop Parse, output.msg, "`n", "`r" {
		thisLine := Trim(A_LoopField)
		if ( !thisLine )
			continue
		if ( InStr(thisLine, "Compressed Hunks of Data") )
			continue
		if ( RegExMatch(thisLine, "i)^Verifying,\s*[\d\.]+%\s+complete") )
			continue
		verifyReport .= thisLine "`n"
	}
	if ( verifyReport )
		threadSendObj.report := verifyReport
}

if ( inStr(recvCmd, "verify") )
	suffx := "verified"
else if ( inStr(recvCmd, "info") )
	suffx := "read info"
else if ( inStr(recvCmd, "extract") )
	suffx := "extracted media"
else if ( inStr(recvCmd, "create") )
	suffx := "created"
else if ( inStr(recvCmd, "addmeta") )
	suffx := "added metadata" 
else if ( inStr(recvCmd, "delmeta") )
	suffx := "deleted metadata" 
else if ( inStr(recvCmd, "copy") )
	suffx := "copied metadata" 
else if ( inStr(recvCmd, "dumpmeta") )
	suffx := "dumped metadata"
else
	suffx := "completed"

threadSendObj.status := "success"
threadSendObj.log := "Successfully " suffx "  -  " recvWorkingTitle
threadSendObj.progressText := "Successfully " suffx "  -  " recvWorkingTitle
threadSendObj.progress := 100

; Calculate file size savings
fileFinishSize := 0
if ( toFile ) {
	outputFile := strReplace(toFile, "`"", "")
	if ( !fileExist(outputFile) ) {
		threadSendObj.status := "error"
		threadSendObj.log := "Error: CHDMAN finished without creating output file"
		threadSendObj.report := "Error: CHDMAN finished without creating output file:`n`n" outputFile "`n"
		if ( output.HasOwnProp("exitcode") )
			threadSendObj.report .= "Exit code: " output.exitcode "`n"
		if ( output.HasOwnProp("msg") && output.msg )
			threadSendObj.report .= "CHDMAN output:`n" output.msg "`n"
		threadSendObj.progressText := "No output file created  -  " recvWorkingTitle
		threadSendObj.progress := 100
		threadSendData()
		thread_cleanupCreatedOutputFolder()
		thread_finishJob()
		ExitApp()
	}
	fileFinishSize := FileGetSize(outputFile) 	; Get new file size
}

if ( instr(recvCmd, "create") > 0 ) { ; If job is compressing a new CHD, add report of file size savings
	threadSendObj.report := "Successfully " suffx "`n`n"
	startSizeForReport := fileStartSize
	; For CUE/GDI/TOC, parser can undercount in edge cases; trust CHDMAN's logical size when it looks more realistic.
	if ( logicalInputSize > 0 && (startSizeForReport <= 0 || startSizeForReport < (logicalInputSize * 0.10)) )
		startSizeForReport := logicalInputSize

	if ( startSizeForReport > 0 ) {
		pcnt := Round((1 - (fileFinishSize / startSizeForReport))*100, 2)
		bytesSaved := startSizeForReport - fileFinishSize
		threadSendObj.report .= "Starting size: " formatBytesReport(startSizeForReport) " - Finished file size: " formatBytesReport(fileFinishSize) "`nTotal size saved: " formatBytesReport(bytesSaved) "  -  Space savings: " pcnt "%"
	}
	else {
		threadSendObj.report .= "Starting size: Unknown - Finished file size: " formatBytesReport(fileFinishSize) "`nTotal size saved: Unknown  -  Space savings: Unknown"
	}
	threadSendObj.fileStartSize  := startSizeForReport
	threadSendObj.fileFinishSize := fileFinishSize
}
else if ( InStr(recvCmd, "extract") > 0 ) {
	outputFolderReport := recvOutputFolder ? recvOutputFolder : splitFilePath(recvToFileFull).dir
	outputFileReport := recvToFileFull ? splitFilePath(recvToFileFull).file : recvWorkingTitle
	threadSendObj.report := "Successfully " suffx "`n`nOutput folder: " outputFolderReport "`nOutput file:   " outputFileReport
}

else
	threadSendObj.report .= "Successfully " suffx "`n`n"

threadSendData()
thread_finishJob()
ExitApp()



; Thread functions
; ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	
; Check for errors after CHDMAN job
; -----------------------------------
thread_checkForErrors(msg)
{
	errorList := ["Error parsing input file", "Error: file already exists", "already exists", "Error opening input file", "Error reading input file", "Unable to open file", "Error writing file"
	, "Error opening parent CHD file", "CHD is uncompressed", "No verification to be done; CHD has no checksum", "Error reading CHD file", "Error creating CHD file"
	, "Error opening CHD file", "Error opening parent CHD file", "Error during compression", "Invalid compressor", "Invalid hunk size"
	, "Unit size is not an even divisor of the hunk size", "Unsupported version", "Error getting info on hunk", "Input start offset greater than input file size"
	, "Can't guess CHS values because there is no input file", "Sector size does not apply"
	, "Error reading audio samples", "Error assembling data for frame", "Invalid size string", "Error opening AVI file", "Error reading AVI frame"
	, "CHD is uncompressed", "CHD has no checksum", "Blank hard disks must be uncompressed", "CHS does not apply when creating a diff from the parent"
	, "Invalid CHS string", "Error reading ident file", "Ident file '", "Template '", "Unable to find hard disk metadata in parent CHD"
	, "Error parsing hard disk metadata in parent CHD", "Data size is not divisible by sector size", "Blank hard drives must specify either a length or a set of CHS values"
	, "Error adding hard disk metadata", "Error adding CD metadata", "Uncompressed is not supported"
	, "Error adding AV metadata", "Error adding AVLD metadata", "Hunk size is not an even multiple or divisor of input hunk size"
	, "Error writing cloned metadata", "Error upgrading CD metadata", "Error writing upgraded CD metadata", "Error writing to file; check disk space"
	, "Unable to recognize CHD file as a CD", "Error writing frame", "Unable to find A/V metadata in the input CHD", "Improperly formatted A/V metadata found"
	, "Frame size does not match hunk size for this CHD", "Error reading hunk", "Error writing samples for hunk", "Error writing video for hunk"
	, "Error reading metadata file", "Error: missing either --valuetext/-vt or --valuefile/-vf parameters", "Error: both --valuetext/-vt or --valuefile/-vf parameters specified; only one permitted"
	, "Error adding metadata", "Error removing metadata:", "Error reading metadata:"]
	
	if ( !msg )
		return ""
	
	for idx, thisErr in errorList {																					
		if ( inStr(msg, thisErr) )					; Check if error from chdman output contains an error string in errorlist array
			return thisErr
	}
	return ""
}



; Get chdman std output, parse it and send it to host
; runCMD() calls this function when it has text data and sends it here
;-------------------------------------------------------------------------
thread_parseCHDMANOutput(data, lineNum, cPID) 													
{ 																	
	;global threadSendObj, threadRecvData, APP_MAIN_NAME, TIMEOUT_SEC
	global
	threadSendObj.chdmanPID := cPID ? cPID : ""

	SetTimer(thread_timeout, TIMEOUT_SEC*1000) 						; Reset timeout timer 
	
	if ( lineNum > 1 ) {
		if ( stPos := inStr(data, "Compressing") ) {
			threadSendObj.status := "compressing"
			,stPos += 13, enPos := inStr(data, "%", false, stPos)		; chdman output: "Compressing, 16.8% complete... (ratio=40.5%)"
			,stPos2 := inStr(data, "(ratio="), enPos2 := inStr(data, "%)",false,0)+2
			,ratio := subStr(data, stPos2, (enPos2-stPos2))
			,threadSendObj.progress := subStr(data, stPos, (enPos-stPos))
			,threadSendObj.progressText := "Compressing -  " ((strLen(threadSendObj.toFile) >= 80)? subStr(threadSendObj.toFile, 1, 66) " ..." : threadSendObj.toFile) (threadSendObj.progress>0 ? "  -  " threadSendObj.progress "% " : "") "  " ratio
		}
		else if ( stPos := inStr(data, "Extracting") ) {
			threadSendObj.status := "extracting"
			,stPos += 12, enPos := inStr(data, "%", false, stPos)		; chdman output: "Extracting, 39.8% complete..."
			,threadSendObj.progress := subStr(data, stPos, (enPos-stPos))
			,threadSendObj.progressText := "Extracting -  " ((strLen(threadSendObj.toFile) >= 90)? subStr(threadSendObj.toFile, 1, 77) " ..." : threadSendObj.toFile) (threadSendObj.progress>0 ? "  -  " threadSendObj.progress "%" : "")
		}
		else if ( stPos := inStr(data, "Verifying") ) {
			threadSendObj.status := "verifying"
			,stPos += 11, enPos := inStr(data, "%", false, stPos)		; chdman output: "Verifying, 39.8% complete..."
			,threadSendObj.progress := subStr(data, stPos, (enPos-stPos))
			,threadSendObj.progressText := "Verifying -  " ((strLen(threadSendObj.fromFile) >= 90)? subStr(threadSendObj.fromFile, 1, 78) " ..." : threadSendObj.fromFile) (threadSendObj.progress>0 ? "  -  " threadSendObj.progress "%" : "") 
		}
		else if ( !inStr(data,"% ") ) { ; Dont capture text that is in the middle of work
			threadSendObj.report := regExReplace(data, "`r|`n", "") "`n"
		}
	}
	; First line is banner/version text; no version gating.
	
	thread_log(data)
	threadSendData()
	return data
}
	

/*
 Send a message to host
---------------------------------------------------------------------------------------
	What we send home:
		threadSendObj.log				-	(string)  General message as to what we are doing
		threadSendObj.status			-	(string)  "started", "done", "error" or "killed" indicating status of job
		threadSendObj.chdmanPID		-	(string)  PID of this chdman
		threadSendObj.report			- 	(string)  Output of chdman and other data to be prsented to user at the end of the job
		threadSendObj.progress		-	(integer) progress percentage
		threadSendObj.progressText	-	(integer) progressbar text description
		-- and all data from host which was previously sent in object
*/
threadSendData(msg:="") 
{
	global

	if ( msg == false )
		return
	
	Sleep(10)
	msg := (msg=="") ? threadSendObj : msg
	if ( (!APP_MAIN_HWND || !WinExist("ahk_id " APP_MAIN_HWND))
		&& IsObject(threadRecvData) && threadRecvData.HasOwnProp("hostPID") && threadRecvData.hostPID )
		APP_MAIN_HWND := WinExist("ahk_class AutoHotkey ahk_pid " threadRecvData.hostPID)
	if ( !APP_MAIN_HWND )
		APP_MAIN_HWND := WinExist(APP_MAIN_NAME)

	msgJSON := jsongo.Stringify(msg)
	targetScript := APP_MAIN_HWND ? "ahk_id " APP_MAIN_HWND : APP_MAIN_NAME
	if ( !sendAppMessage(msgJSON, targetScript) && APP_MAIN_HWND )
		sendAppMessage(msgJSON, APP_MAIN_NAME)										; Fallback in case HWND changed
	threadSendObj.log := ""
	threadSendObj.report := ""
	threadSendObj.status := ""
	Sleep(10)
}


/*
Recieve messages from host
---------------------------------------------------------------------------------------	
		What we recieve from host:
		q.PID			 - (string)  PID of (this) thread which starts chdman.exe
		q.idx		      - (string)  Job Number in queue
		q.cmd 			 - (string)  The command for chdman to run (ie 'extractcd', 'createhd', 'verify', 'info')
		q.cmdOpts 		 - (string)  The options (with parameters) to pass along to chdman
		q.workingDir 	 - (string)  The input working directory
		q.fromFile 		 - (string)  Input filename without path
		q.fromFileExt	 - (string)  Input file extension
		q.fromFileNoExt  - (string)  Input filename without path or extension
		q.fromFileFull 	 - (string)  Input filename with full path and extension
		q.outputFolder 	 - (string)  The output folder where files will be saved
		q.toFile 		 - (string)  Output filename without path
		q.toFileExt		 - (string)  Output file extension
		q.toFileNoExt 	 - (string)  Output filename without path or extension
		q.toFileFull 	 - (string)  Output filename with full path and extension
		q.fileDeleteList - (array)   Files set to be deleted after job has completed
		q.hostPID		 - (string)  PID of main program
		q.id 			 - (number)  Unique job id
*/
thread_receiveData(wParam, lParam) 
{
	global

	threadRecvData := thread_jsonToLegacyObj(jsongo.Parse(StrGet(NumGet(lParam + 2*A_PtrSize, 0, "UPtr"),, "utf-8")))
	if ( !IsObject(threadRecvData) )
		return 0
	
	if ( threadRecvData.HasOwnProp("KILLPROCESS") && threadRecvData.KILLPROCESS == "true" ) {
		recvWorkingTitle := threadRecvData.HasOwnProp("workingTitle") ? threadRecvData.workingTitle : "Unknown job"
		jobKeepIncomplete := IsObject(threadJobData) && threadJobData.HasOwnProp("keepIncomplete") && threadJobData.keepIncomplete
		jobToFileFull := (IsObject(threadJobData) && threadJobData.HasOwnProp("toFileFull")) ? threadJobData.toFileFull : ""
		recvIdx := threadRecvData.HasOwnProp("idx") ? threadRecvData.idx : "?"
		
		threadSendObj.log := "Attempting to cancel " . recvWorkingTitle
		threadSendObj.progressText := "Cancelling -  " recvWorkingTitle
		threadSendObj.progress := 0
		threadSendData()
		
		thread_log("`nThread cancel signal receieved`n")

		cancelPid := 0
		if ( threadRecvData.HasOwnProp("chdmanPID") && threadRecvData.chdmanPID )
			cancelPid := threadRecvData.chdmanPID
		else if ( IsObject(RUNCMD_STATE) && RUNCMD_STATE.HasOwnProp("PID") && RUNCMD_STATE.PID )
			cancelPid := RUNCMD_STATE.PID

		if ( cancelPid ) {
			try ProcessClose(cancelPid)
			catch {
			}
			Sleep(500)
		}
		if ( cancelPid && ProcessExist(cancelPid) ) { ; Process still exists
		
			threadSendObj.log := "Couldn't cancel " recvWorkingTitle " - Error closing job"
			threadSendObj.progressText := "Couldn't cancel -  " recvWorkingTitle
			threadSendObj.progress := 100
			threadSendObj.report := "`nCancelling of job was unsuccessful`n"
			threadSendData()
			
			thread_log("`n`nJob couldn't be cancelled!`n")
		}
		else { 
			threadSendObj.status := "cancelled"
			threadSendObj.report := "Job cancelled by user`n`n"
			threadSendData()
			
			managedOutputFolder := !jobKeepIncomplete ? thread_getManagedOutputSubdir() : ""
			if ( managedOutputFolder ) {
				if ( folderDelete(managedOutputFolder, 5, 50, 1) ) {
					threadJobData.outputFolderCreatedByJob := false
					threadSendObj.log := "Deleted created output folder '" managedOutputFolder "'"
					threadSendObj.report := "Deleted created output folder:`n    " managedOutputFolder "`n"
				}
				else {
					threadSendObj.log := "Error deleting created output folder '" managedOutputFolder "'"
					threadSendObj.report := threadSendObj.log "`n"
				}
				threadSendObj.progress := 100
				threadSendData()
			}
			else if ( !jobKeepIncomplete
				&& jobToFileFull
				&& fileExist(jobToFileFull) ) {						; Delete incomplete output files if asked to keep 
				delFiles := deleteFilesReturnList(jobToFileFull)
				threadSendObj.log := delFiles != "" ? "Deleted incomplete file(s): " regExReplace(delFiles, " ,$") : "Error deleting incomplete file(s)!"
				threadSendObj.report := delFiles != "" ? thread_formatReportList("Deleted incomplete file(s):", delFiles) : threadSendObj.log "`n"
				threadSendObj.progress := 100
				threadSendData()
			}
			thread_cleanupCreatedOutputFolder()
			
			threadSendObj.log := "Job " recvIdx " cancelled by user"
			threadSendObj.progressText := "Cancelled -  " recvWorkingTitle
			threadSendObj.progress := 100
			threadSendObj.report := ""
			threadSendData()										
			
			thread_finishJob()
			
			thread_log("`nJob cancelled by user`n")
			ExitApp()
		}
	}
	return 1
}

thread_jsonToLegacyObj(val)
{
	if ( val is Map ) {
		obj := {}
		for key, item in val
			obj.%key% := thread_jsonToLegacyObj(item)
		return obj
	}
	if ( val is Array ) {
		arr := []
		for _, item in val
			arr.Push(thread_jsonToLegacyObj(item))
		return arr
	}
	return val
}
	

; timer is refreshed on every call of thread_parseCHDMANOutput - if it lands here we assume chdman has timed out
; ------------------------------------------------------------------------------------------
thread_timeout()
{
	global threadRecvData, threadSendObj, RUNCMD_STATE
	recvWorkingTitle := (IsObject(threadRecvData) && threadRecvData.HasOwnProp("workingTitle")) ? threadRecvData.workingTitle : "Unknown job"

	runningPid := 0
	if ( threadSendObj.HasOwnProp("chdmanPID") && threadSendObj.chdmanPID )
		runningPid := threadSendObj.chdmanPID
	else if ( IsObject(RUNCMD_STATE) && RUNCMD_STATE.HasOwnProp("PID") && RUNCMD_STATE.PID )
		runningPid := RUNCMD_STATE.PID

	; Avoid false timeouts while CHDMAN is still alive but temporarily silent.
	if ( runningPid && ProcessExist(runningPid) ) {
		SetTimer(thread_timeout, TIMEOUT_SEC*1000)
		return
	}
	
	threadSendObj.status := "error"
	threadSendObj.progressText := "Error  -  CHDMAN timeout " recvWorkingTitle
	threadSendObj.progress := 100
	threadSendObj.log := "Error: Job failed - CHDMAN timed out"
	threadSendObj.report := "Error: Job failed - CHDMAN timed out`n`n"
	threadSendData()
	thread_cleanupCreatedOutputFolder()
	
	thread_finishJob()				; contains threadSendData()
	ExitApp()
}

thread_heartbeat()
{
	global threadJobWatch, threadSendObj, threadRecvData, RUNCMD_STATE
	recvWorkingTitle := (IsObject(threadRecvData) && threadRecvData.HasOwnProp("workingTitle")) ? threadRecvData.workingTitle : "Unknown job"

	if ( !IsObject(threadJobWatch) || !threadJobWatch.active )
		return

	runningPid := 0
	if ( threadSendObj.HasOwnProp("chdmanPID") && threadSendObj.chdmanPID )
		runningPid := threadSendObj.chdmanPID
	else if ( IsObject(RUNCMD_STATE) && RUNCMD_STATE.HasOwnProp("PID") && RUNCMD_STATE.PID )
		runningPid := RUNCMD_STATE.PID

	if ( !runningPid || !ProcessExist(runningPid) )
		return

	if ( InStr(threadJobWatch.cmd, "create") && threadJobWatch.outputFile && FileExist(threadJobWatch.outputFile) ) {
		currSize := FileGetSize(threadJobWatch.outputFile)
		if ( threadJobWatch.startSize > 0 ) {
			pct := Floor((currSize / threadJobWatch.startSize) * 100)
			if ( pct < 1 && currSize > 0 )
				pct := 1
			if ( pct > 99 )
				pct := 99

			if ( pct != threadJobWatch.lastPct ) {
				threadJobWatch.lastPct := pct
				threadSendObj.progress := pct
				threadSendObj.progressText := "Working -  " recvWorkingTitle "  -  " pct "% (" formatBytes(currSize) ")"
				threadSendData()
				thread_log("Working... " pct "% (" formatBytes(currSize) ")`n")
				return
			}
		}
	}

	if ( (A_TickCount - threadJobWatch.lastTick) >= 5000 ) {
		threadJobWatch.lastTick := A_TickCount
		if ( !threadSendObj.HasOwnProp("progress") || threadSendObj.progress < 1 )
			threadSendObj.progress := 1
		threadSendObj.progressText := "Working -  " recvWorkingTitle
		threadSendData()
	}
}
; Delete input files after a successful completeion of CHDMAN (if specified in options)
; -----------------------------------------------------------------------------------------
thread_deleteFiles(delfile) 
{
	global threadSendObj, threadRecvData
	
	deleteTheseFiles := getFilesFromCUEGDITOC(delfile)					; Get files to be deleted
	if ( deleteTheseFiles.Length == 0 )
		return false
	
	log := "", errLog := ""
	for idx, thisFile in deleteTheseFiles {			
		if ( deleteFileWithRetry(thisFile, 3, 25) )
			log .= "'" splitFilePath(thisFile).file "'`n"
		else
			errlog .= "'" splitFilePath(thisFile).file "'`n"
	}
	
	if ( log )
		threadSendObj.log := "Deleted Files: " RegExReplace(RegExReplace(log, "`n+$"), "`n", ", ")
	if ( errLog )
		threadSendObj.log := (log ? "Deleted Files: " RegExReplace(RegExReplace(log, "`n+$"), "`n", ", ") "`n" : "") "Error deleting: " RegExReplace(RegExReplace(errLog, "`n+$"), "`n", ", ")
	
	threadSendObj.report := ""
	if ( log )
		threadSendObj.report .= thread_formatReportList("Deleted Files:", log, false)
	if ( errLog )
		threadSendObj.report .= (threadSendObj.report ? "`n" : "") thread_formatReportList("Error deleting:", errLog, false)
	threadSendObj.progress := 100
	threadSendData()
	
	thread_log(threadSendObj.log "`n")
}


thread_deleteDir(dir, delFull:=0) 
{
	if ( !delFull && !dllCall("Shlwapi\PathIsDirectoryEmpty", "Str", dir) )
		threadSendObj.log := "Error deleting directory '" dir "' - Not empty"
	else
		threadSendObj.log := folderDelete(dir, 5, 50, delFull) ? "Deleted directory '" dir "'" : "Error deleting directory '" dir "'"
	
	threadSendObj.report := threadSendObj.log "`n"
	threadSendObj.progress := 100
	threadSendData()
	
	thread_log(threadSendObj.log "`n")
}

thread_cleanupCreatedOutputFolder()
{
	global threadRecvData, threadJobData
	outputFolder := ""

	if ( !IsObject(threadJobData)
		|| (threadJobData.HasOwnProp("keepIncomplete") && threadJobData.keepIncomplete)
		|| !threadJobData.HasOwnProp("outputFolder")
		|| !threadJobData.outputFolder )
		return false

	outputFolder := thread_getManagedOutputSubdir()
	if ( !outputFolder || FileExist(outputFolder) != "D" )
		return false

	if ( folderDelete(outputFolder, 5, 50, 1) ) {
		threadJobData.outputFolderCreatedByJob := false
		if ( IsObject(threadRecvData) )
			threadRecvData.outputFolderCreatedByJob := false
		thread_log("Deleted created output directory '" outputFolder "'`n")
		return true
	}

	if ( DllCall("Shlwapi\PathIsDirectoryEmpty", "Str", outputFolder)
		&& folderDelete(outputFolder, 5, 50, 0) ) {
		threadJobData.outputFolderCreatedByJob := false
		if ( IsObject(threadRecvData) )
			threadRecvData.outputFolderCreatedByJob := false
		thread_log("Deleted created output directory '" outputFolder "'`n")
		return true
	}
	return false
}

thread_getManagedOutputSubdir()
{
	global threadJobData
	outputFolder := ""
	workingDir := ""
	fromFileFull := ""
	cmd := ""

	if ( !IsObject(threadJobData)
		|| !threadJobData.HasOwnProp("outputFolder")
		|| !threadJobData.outputFolder
		|| !threadJobData.HasOwnProp("outputFolderIsPerJob")
		|| !threadJobData.outputFolderIsPerJob
		|| !threadJobData.HasOwnProp("toFileNoExt")
		|| !threadJobData.toFileNoExt )
		return ""

	outputFolder := RegExReplace(threadJobData.outputFolder, "\\+$")
	if ( FileExist(outputFolder) != "D" )
		return ""

	workingDir := threadJobData.HasOwnProp("workingDir") ? RegExReplace(threadJobData.workingDir, "\\+$") : ""
	fromFileFull := threadJobData.HasOwnProp("fromFileFull") ? threadJobData.fromFileFull : ""
	cmd := threadJobData.HasOwnProp("cmd") ? threadJobData.cmd : ""

	if ( splitFilePath(outputFolder).file != threadJobData.toFileNoExt )
		return ""
	if ( workingDir && outputFolder == workingDir )
		return ""
	if ( fromFileFull && InStr(fromFileFull, outputFolder "\") )
		return ""

	if ( RegExMatch(cmd, "^extract") )
		return outputFolder
	if ( RegExMatch(cmd, "^create")
		&& threadJobData.HasOwnProp("outputFolderCreatedByJob")
		&& threadJobData.outputFolderCreatedByJob )
		return outputFolder

	return ""
}

thread_formatReportList(title, listText, commaDelimited:=true)
{
	listBody := commaDelimited ? RegExReplace(RegExReplace(listText, ",\s*$"), ",\s*", "`n") : RegExReplace(listText, "`n+$")
	listBody := RegExReplace(listBody, "m)^(?!\s*$)", "    ")
	return title "`n" listBody "`n"
}


; Send output to console
;---------------------------------------------------------------------------------------
thread_log(newMsg, concat:=true) 
{
	global threadConsole
	
	if ( !isObject(threadConsole) )
		return

	if ( threadConsole.getConsoleHWND() )
		threadConsole.write(newMsg)
}



; Extract an archive file
;http://www.autohotkey.com/forum/viewtopic.php?p=402574
; -----------------------------------------------------
thread_extractArchive(file, dir)
{
	ext := StrLower(splitFilePath(file).ext)
	if ( ext == "zip" )
		return thread_unzip(file, dir)
	if ( ext == "7z" )
		return thread_extract7z(file, dir)
	return false
}

; Unzip a zip file
; ----------------
thread_unzip(file, dir)
{
    global threadRecvData, threadSendObj
	recvFromFile := (IsObject(threadRecvData) && threadRecvData.HasOwnProp("fromFile")) ? threadRecvData.fromFile : "file"
	
	try {
		psh  := ComObject("Shell.Application")
		srcNs := psh.Namespace(file)
		dstNs := psh.Namespace(dir)
		if ( !IsObject(srcNs) || !IsObject(dstNs) )
			return false

		srcItems := srcNs.items
		zipped := srcItems.count
		if ( zipped < 1 )
			return false

		dstNs.CopyHere(srcItems, 4|16)

		startTick := A_TickCount
		lastCount := -1
		stableTicks := 0
		lastShown := -1
		lastSentTick := 0
		pulse := 0

		loop {
			Sleep(120)
			unzipped := dstNs.items().count

			if ( unzipped = lastCount )
				stableTicks++
			else {
				lastCount := unzipped
				stableTicks := 0
			}

			pct := Floor((unzipped / zipped) * 100)
			if ( pct < 1 ) {
				pulse := Mod(pulse + 4, 30)
				showPct := 1 + pulse
			}
			else
				showPct := pct

			if ( showPct > 99 )
				showPct := 99

			if ( showPct != lastShown || (A_TickCount - lastSentTick) >= 500 ) {
				threadSendObj.status := "unzipping"
				threadSendObj.progress := showPct
				threadSendObj.progressText := "Unzipping  -  " recvFromFile "  -  " showPct "%"
				threadSendData()
				lastShown := showPct
				lastSentTick := A_TickCount
			}

			; Consider extraction done once item count reached and stayed stable briefly.
			if ( unzipped >= zipped && stableTicks >= 6 )
				break

			; Hard stop to avoid hanging forever on shell extraction edge-cases.
			if ( (A_TickCount - startTick) > 300000 )
				break
		}
		
		loop Files regExReplace(dir, "\\$") "\*.*", "FR"
		{
			zipfile := splitFilePath(A_LoopFileFullPath)
			zipFileExtLower := StrLower(zipfile.ext)
			if ( inArray(zipFileExtLower, threadRecvData.inputFileTypes) && !isArchiveExt(zipFileExtLower) )
				return zipfile			; Use only the first file found in the zip temp dir
		}
		return false
	
	}
	catch as e
		return false
}

thread_extract7z(file, dir)
{
	global threadRecvData, SEVENZIP_EXE

	if ( !SEVENZIP_EXE || !FileExist(SEVENZIP_EXE) )
		return false

	cmd := "`"" SEVENZIP_EXE "`" x -y -bd -bso0 -bsp0 -o`"" dir "`" -- `"" file "`""
	output := runCMD(cmd, "", "UTF-8")
	if ( !output || !output.HasOwnProp("exitcode") || output.exitcode != 0 )
		return false

	loop Files regExReplace(dir, "\\$") "\*.*", "FR"
	{
		zipfile := splitFilePath(A_LoopFileFullPath)
		zipFileExtLower := StrLower(zipfile.ext)
		if ( inArray(zipFileExtLower, threadRecvData.inputFileTypes) && !isArchiveExt(zipFileExtLower) )
			return zipfile
	}
	return false
}




; Finish the job
; ----------------
thread_finishJob() 
{
	global

	finalStatus := threadSendObj.HasOwnProp("status") ? threadSendObj.status : ""
	if ( finalStatus == "error" || finalStatus == "cancelled" || finalStatus == "fileExists" || finalStatus == "halted" ) {
		Sleep(100)
		thread_cleanupCreatedOutputFolder()
	}

	Sleep(10)
	threadSendObj.status := "finished"
	threadSendObj.progress := 100
	thread_log(threadSendObj.log? threadSendObj.log "`nFinished!":"")
	threadSendData()
	
	if ( SHOW_JOB_CONSOLE == "yes" )	
		Sleep(WAIT_TIME_CONSOLE_SEC*1000)	; Wait x seconds or until user closes window
	ExitApp()
}

#Warn VarUnset


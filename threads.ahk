onMessage(0x4a, "thread_receiveData")
onExit("thread_Quit", 1)

gui 1:show, hide, % runAppName
menu tray, noIcon

thread_log("Started ... ")

recvData := {}, sendData := {}, mainAppHWND := winExist(mainAppName), verifedAfterFromFile:=false

console := new Console()
console.setConsoleTitle(runAppNameConsole)
;winSet, Style, ^0x80000 , % "ahk_id " console.getConsoleHWND() ; Remove close button on window
if ( a_args[2] <> "console" ) 												; Hide console
	winHide , % "ahk_id " console.getConsoleHWND()

thread_log("Handshaking with " mainAppName "... ")
while ( !recvData.cmd && !recvData.idx ) {
	thread_log(".")
	sleep 10
	if ( a_index > 1000 ) {
		thread_log("Error handshaking!`n")
		return
	}
}
thread_log("OK!`n"
		. "Starting a " stringUpper(recvData.cmd) " job`n`n"
		. "Working job file: " recvData.workingTitle "`n"
		. "Working directory: " recvData.workingDir "`n")

mergeObj(recvData, sendData)															; Assign recvData to sendData as we will be sending the same info back and forth

sendData.status := "started"
sendData.report := stringUpper(recvData.cmd) " - " recvData.workingTitle "`n" drawLine(77) "`n"
sendData.pid := dllCall("GetCurrentProcessId")
sendData.progress := 0
sendData.log := "Starting " stringUpper(recvData.cmd) " job - " recvData.workingTitle
sendData.progressText := "Starting job  -  " recvData.workingTitle
thread_sendData()	

if ( fileExist(recvData.outputFolder) <> "D" ) {
	fileCreateDir, % recvData.outputFolder
	thread_log("Created directory: " recvData.outputFolder "`n")
}

if ( recvData.fromFileExt == "zip" ) {
	tempZipDirectory := mainTempDir "\" recvData.fromFileNoExt
	folderDelete(tempZipDirectory, 5, 50, 1)
	fileCreateDir, % tempZipDirectory
	
	if ( fileExist(tempZipDirectory) == "D" ) {
		thread_log("Unzipping " recvData.fromFileFull "`nUnzipping to: " tempZipDirectory)
		
		sendData.status := "unzipping"
		sendData.log := "Unzipping " recvData.fromFile " to " tempZipDirectory
		sendData.progressText := "Unzipping  -  " recvData.fromFile
		sendData.progress := 0
		thread_sendData()
		
		if ( unzip(recvData.fromFileFull, tempZipDirectory) ) {
			Loop, Files, % tempZipDirectory "\*.*", FR
			{
				f := splitPath(a_LoopFileLongPath)
				if ( inArray(f.ext, recvData.inputFileTypes) ) {
					recvData.fromFileFull := f.full
					
					fromFile := f.full ? """" f.full """"  : ""
					
					sendData.status := "unzipping"
					sendData.log := "Unzipped " f.full " successfully"
					sendData.report := "Unzipped " f.full " successfully`n"
					sendData.progressText := "Unzipped successfully  -  " f.file
					sendData.progress := 0
					thread_sendData()
					break  								; Use only the first file found in the zip temp dir
				}
			}
			if ( !fromFile )
				error := ["Error finding unzipped files", "Error finding unzipped files"]
		}
		else 
			error := ["Error unzipping file '" recvData.fromFileFull "'", "Error unzipping file  -  " recvData.fromFileFull]
	}
	else
		error := ["Error creating temporary directory '" mainTempDir "\" recvData.fromFileNoExt "'", "Error creating temp directory"]
	
	if ( error ) {
		sendData.status := "error"
		sendData.log := error[1]
		sendData.report := "`n" error[1] "`n"
		sendData.progressText := error[2] "  -  " recvData.workingTitle
		sendData.progress := 100
		thread_sendData()

		thread_finishJob() ; Deletes files 
		exitApp
	}
}
else
	fromFile := recvData.fromFileFull ? """" recvData.fromFileFull """" : ""

toFile := recvData.toFileFull ? """" recvData.toFileFull """" : ""

cmdLine := chdmanLocation . " " . recvData.cmd . recvData.cmdOpts . " -v" . (fromFile ? " -i " fromFile : "") . (toFile ? " -o " toFile : "")
thread_log("`nCommand line: " cmdLine "`n`n") 

setTimer, thread_timeout, % (timeoutSec*1000) 						; Set timeout timer
output := runCMD(cmdLine, recvData.workingDir, "CP0", "thread_parseCHDMANOutput")
setTimer, thread_timeout, off

if ( tempZipDirectory )
	thread_deleteDir(tempZipDirectory, 1) ; Always delete temp directory
	
rtnError := thread_checkForErrors(output.msg)

; CHDMAN was not successfull
; Errors were detected
; -----------------------
if ( (rtnError && rtnError <> "file already exists") || (inStr(rtnError, "file already exists") && !inStr(recvData.cmdOpts, "-f")) ) {
	if ( !recvData.keepIncomplete && !inStr(rtnError, "file already exists")  )			; Delete incomplete output files, but dont delete "incomplete" output file if the error is that the file exists
		thread_deleteIncompleteFiles(recvData.toFileFull)
	
	sendData.status := "error"
	sendData.log := rtnError
	sendData.report := "`n" (inStr(sendData.log, "Error") ? "" : "Error: ") sendData.log "`n"
	sendData.progressText := regExReplace(sendData.log, "`n|`r", "") "  -  " recvData.workingTitle
	sendData.progress := 100
	thread_sendData()

	thread_finishJob()
	exitApp
}


; CHDMAN was successfull
; No errors were detected
; -----------------------
if ( recvData.deleteInputFiles ) 				; Delete input files if requested
	thread_deleteFiles(recvData.fromFileFull)

if ( recvData.deleteInputDir )					; Delete input folder only if its not empty
	thread_deleteDir(recvData.workingDir)	

if ( inStr(recvData.cmd, "verify") )
	suffx := "verified"
else if ( inStr(recvData.cmd, "extract") )
	suffx := "extracted media"
else if ( inStr(recvData.cmd, "create") )
	suffx := "created"
else if ( inStr(recvData.cmd, "addmeta") )
	suffx := "added metadata" 
else if ( inStr(recvData.cmd, "delmeta") )
	suffx := "deleted metadata" 
else if ( inStr(recvData.cmd, "copy") )
	suffx := "copied metadata" 
else if ( inStr(recvData.cmd, "dumpmeta") )
	suffx := "dumped metadata"
	
sendData.status := "success"
sendData.log := "Successfully " suffx "  -  " recvData.workingTitle
sendData.report := "`nSuccessfully " suffx "`n"
sendData.progressText := "Successfully " suffx "  -  " recvData.workingTitle
sendData.progress := 100
thread_sendData()

thread_finishJob() ; Finish job
exitApp
	


; Thread functions
; ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;	
	
; Check for errors after CHDMAN job
; -----------------------------------
thread_checkForErrors(msg)
{
	errorList := ["Error parsing input file", "Error: file already exists", "Error opening input file", "Error reading input file", "Unable to open file", "Error writing file"
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
	global sendData, recvData, chdmanVerArray, mainAppName, timeoutSec
	sendData.chdmanPID := cPID ? cPID : ""

	setTimer, thread_timeout, % (timeoutSec*1000) 				; Reset timeout timer 
	
	if ( lineNum > 1 ) {
		if ( stPos := inStr(data, "Compressing") ) {
			sendData.status := "compressing"
			,stPos += 13, enPos := inStr(data, "%", false, stPos)		; chdman output: "Compressing, 16.8% complete... (ratio=40.5%)"
			,stPos2 := inStr(data, "(ratio="), enPos2 := inStr(data, "%)",false,0)+2
			,ratio := subStr(data, stPos2, (enPos2-stPos2))
			,sendData.progress := subStr(data, stPos, (enPos-stPos))
			,sendData.progressText := "Compressing -  " ((strLen(sendData.toFile) >= 80)? subStr(sendData.toFile, 1, 66) " ..." : sendData.toFile) (sendData.progress>0 ? "  -  " sendData.progress "% " : "") "  " ratio
		}
		else if ( stPos := inStr(data, "Extracting") ) {
			sendData.status := "extracting"
			,stPos += 12, enPos := inStr(data, "%", false, stPos)		; chdman output: "Extracting, 39.8% complete..."
			,sendData.progress := subStr(data, stPos, (enPos-stPos))
			,sendData.progressText := "Extracting -  " ((strLen(sendData.toFile) >= 90)? subStr(sendData.toFile, 1, 77) " ..." : sendData.toFile) (sendData.progress>0 ? "  -  " sendData.progress "%" : "")
		}
		else if ( stPos := inStr(data, "Verifying") ) {
			sendData.status := "verifying"
			,stPos += 11, enPos := inStr(data, "%", false, stPos)			; chdman output: "Verifying, 39.8% complete..."
			,sendData.progress := subStr(data, stPos, (enPos-stPos))
			,sendData.progressText := "Verifying -  " ((strLen(sendData.fromFile) >= 90)? subStr(sendData.fromFile, 1, 78) " ..." : sendData.fromFile) (sendData.progress>0 ? "  -  " sendData.progress "%" : "") 
		}
		else if ( !inStr(data,"% ") ) { ; Dont capture text that is in the middle of work
			sendData.report := regExReplace(data, "`r|`n", "") "`n"
		}
	}
	else {
		enPos := inStr(data, "(", false, stPos)
		chdmanVer := trim(subStr(data, 53, (enPos-53)))
		
		if ( !inArray(chdmanVer, chdmanVerArray) )  {
			sendData.status := "error"
			sendData.log := "Error: Wrong CHDMAN version " chdmanVer "`n - Supported versions of CHDMAN are: " arrayToString(chdmanVerArray)
			sendData.report := "Wrong CHDMAN version supplied [" chdmanVer "]`nSupported versions of CHDMAN are: " arrayToString(chdmanVerArray) "`n`nJob cancelled.`n"
			sendData.progressText := "Error - Wrong CHDMAN version -  " recvData.workingTitle
			sendData.progress := 100
			thread_log(sendData.log "`n")
			thread_sendData()
			
			thread_finishJob()
			exitApp
		}
	}
	
	thread_log(data)
	thread_sendData()

	return data
}
	

/*
 Send a message to host
---------------------------------------------------------------------------------------
	What we send home:
		sendData.log				-	(string)  General message as to what we are doing
		sendData.status			-	(string)  "started", "done", "error" or "killed" indicating status of job
		sendData.chdmanPID		-	(string)  PID of this chdman
		sendData.report			- 	(string)  Output of chdman and other data to be prsented to user at the end of the job
		sendData.progress		-	(integer) progress percentage
		sendData.progressText	-	(integer) progressbar text description
		-- and all data from host which was previously sent in object
*/
thread_sendData(msg:="") 
{
	global recvData, sendData,  mainAppName, mainAppHWND

	if ( msg == false )
		return
	msg := (msg=="") ? sendData : msg
	sendAppMessage(toJSON(msg), mainAppName " ahk_id " mainAppHWND)										; Send back the data we've recieved plus any other new info
	sendData.log := ""
	sendData.status := ""
	sendData.report := ""
	sendData.progress := ""
	sendData.progressText := ""
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
thread_receiveData(wParam, ByRef lParam) 
{
	global recvData
	
	stringAddress := numGet(lParam + 2*A_PtrSize) 
	recvData := fromJSON(strGet(stringAddress,, "utf-8"))
	
	if ( recvData.kill == "true" )									; User has requested this job to be cancelled
		thread_killProcess()
}
	

	
; Kill process
; -----------------------------------------------------------------------------------------
thread_killProcess() 
{
	global
	critical

	a_Args.runCMD.PID := 0
	processPIDClose(sendData.chdmanPID, 10, 250)

	sendData.status := "killed"
	sendData.progressText := "Cancelled -  " recvData.workingTitle
	sendData.progress := 100
	sendData.log := "CHDMAN process (PID " sendData.chdmanPID ") killed or cancelled by user"
	thread_log(sendData.log "`n")
	thread_sendData()

	if ( !recvData.keepIncomplete )
		thread_deleteIncompleteFiles(recvData.toFileFull) 	; Will add to log and report
	
	sendData.report := "`nError: Job killed or cancelled by user`n"
	thread_sendData()
	
	thread_finishJob()
	exitApp
	
}

; timer is refreshed on every call of thread_parseCHDMANOutput - if it lands here we assume chdman has timed out
; ------------------------------------------------------------------------------------------
thread_timeout()
{
	global recvData, sendData
	
	sendData.status := "error"
	sendData.progressText := "Error  -  CHDMAN timeout " recvData.workingTitle
	sendData.progress := 100
	sendData.log := "Error: Job failed - CHDMAN timed out"
	sendData.report := "`nError: Job failed - CHDMAN timed out" "`n"
	thread_sendData()
	
	thread_finishJob() ; contains thread_sendData()
	exitApp	
}



; Delete files after an unsuccessful completeion of CHDMAN
; -----------------------------------------------------------------------------------------
thread_deleteIncompleteFiles(file) 
{
	global recvData, sendData

	if ( !fileExist(file) )
		return false
		
	deleteTheseFiles := getFilesFromCUEGDITOC(file)					; Get files to be deleted
	
	sendData.status := "deleting"
	delFiles := ""
	for idx, thisFile in deleteTheseFiles {
		delFiles .= fileDelete(thisFile, 5, 50) ? thisFile ", " : ""
	}
	sendData.log := delFiles ? "Deleted incomplete file(s): " regExReplace(delFiles, " ,$") : "Error deleting incomplete file(s)!"
	sendData.report := sendData.log "`n"
	sendData.progress := 100
	thread_log(sendData.log "`n")
	thread_sendData()
}
	
	
	
; Delete input files after a successful completeion of CHDMAN (if specified in options)
; -----------------------------------------------------------------------------------------
thread_deleteFiles(delfile) 
{
	global sendData, recvData
	
	deleteTheseFiles := getFilesFromCUEGDITOC(delfile)					; Get files to be deleted
	if ( deleteTheseFiles.length() == 0 )
		return false
	
	log := "", errLog := ""
	for idx, thisFile in deleteTheseFiles {			
		if ( fileDelete(thisFile, 5) )
			log .= "'" splitPath(thisFile).file "', "
		else
			errlog .= "'" splitPath(thisFile).file "', "
	}
	
	sendData.status := "deleting"
	sendData.progress := 100
	sendData.log := ""
	if ( log )
		sendData.log := "Deleted Files: " regExReplace(log, ", $")	
	if ( errLog )
		sendData.log := (log ? log "`n" : "") "Error deleting: " regExReplace(errLog, ", $")									; Remove trailing comma
	
	sendData.report := sendData.log "`n"
	thread_log(sendData.log "`n")
	thread_sendData()
}


thread_deleteDir(dir, delFull:=0) 
{
	if ( !delFull && !dllCall("Shlwapi\PathIsDirectoryEmpty", "Str", dir) )
		sendData.log := "Error deleting directory '" dir "' - Not empty"
	else
		sendData.log := folderDelete(dir, 5, 50, delFull) ? "Deleted directory '" dir "'" : "Error deleting directory '" dir "'"
	sendData.report := sendData.log "`n"
	
	sendData.progress := 100
	thread_sendData()
	
	thread_log(sendData.log "`n")
}



; Finish the job
; ----------------
thread_finishJob() 
{
	global sendData, recvData
	
	sleep 150
	
	sendData.status := "finished"
	sendData.report .= "`n`n"
	thread_log(sendData.log? sendData.log "`nFinished!":"")
	thread_sendData()
}



; Send output to console
;---------------------------------------------------------------------------------------
thread_log(newMsg, concat=true) 
{
	global console
	
	if ( console.getConsoleHWND() )
		console.write(newMsg)
}
	
	
; Quit the thread
;---------------------------------------------------------------------------------------
thread_quit() 
{
	global console, waitTimeConsoleSec
	if ( console.getConsoleHWND() )	
		sleep waitTimeConsoleSec*1000	; Wait x seconds or until user closes window

}
	
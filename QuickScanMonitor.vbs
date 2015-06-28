'*******************************************************************************
'* File:           QuickScanMonitor
'*
'* Purpose:        AUTOMATIC QUICKSCAN & ARTESIA/TEAMS FILE PROCESSING AND IMPORTING
'* Version:        1.0 (18 October 2011)
'*
'* Author:		   	 Jon Normington
'* Requirements:   > Windows XP
'*                 Windows Script Host 5.6 - CSCRIPT.EXE_OR_WSCRIPT.EXE
'*******************************************************************************

'Continue if errors occur
On Error Resume Next

Dim currentChecks, newFilesEmailBody, emailBodyInfo, monitorResult, fileRename, file, unpfile
Dim WshShell, objFSH, filesToProcessList, newFilesGenerated, objWMIService

Const checksBeforeMail = 50 'total seconds = (checksBeforeMail * timeToSleep) / 1000
Const timeToSleep = 5000 'Sleep for 5 seconds before checking the process again
Const logDirectory = "E:\InvoiceUploads\logs\" 'Logs Directory
Const toBeProcessDirectory = "E:\APScannedInvoices\" 'London pickup folder location dropped by the network scanner
Const toBeProcessDirectorySweden = "E:\APScannedInvoicesSE\" 'Sweden pickup folder location dropped by the network scanner
Const processingDirectory = "E:\InvoiceUploads\working_area\" ' Working directory for both London/Sweden
Const readyToUploadDirectory = "E:\InvoiceUploads\scans" 'London directory that the client PC running Artesia plugin watches
Const readyToUploadDirectorySweden = "E:\InvoiceUploads\swedenscans" 'Sweden directory that the client PC running Artesia plugin watches
Const strComputer = "." 'Means local PC - used to get the running tasks on a PC to monitor QuickScan Pro
Const processName = "Quickscn.exe" 'Process name of the task we monitor
Const quickCmd = """C:\Program Files\EMC Captiva\QuickScan\quickscn.exe""" 'To call QuickScanPro from command line - note the extra quotes required.
Const quickScanArgments = " /scan profile=TEAMS_Import - FileImport" 'London scan template to call the the quickCmd
Const quickScanArgmentsSweden = " /scan profile=TEAMS_Import Sweden - FileImport" 'Sweden scan template to call the the quickCmd
Const emailFrom = "<emailFrom>" 'Email from server.
Const emailTo = "<emailTo>" 'Email to if any directories required are missing.
Const smtp = "smtp.domain.com"
Const archiveDir = "<archiveDirectory>"

'Even if we error cause no directory exist it is
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSH = WScript.CreateObject("Scripting.FileSystemObject")
Set unProcessedFileList = objFSH.GetFolder(processingDirectory).Files
Set filesToProcessList = objFSH.GetFolder(toBeProcessDirectory).Files

'Sweden files to be picked up
Set filesToProcessListSweden = objFSH.GetFolder(toBeProcessDirectorySweden).Files
Set logFileWriter = objFSH.OpenTextFile(logDirectory & "invoiceAuto_" & StrCleaner(Date, "-") & ".log", 8, True)
LogEvent "Script Started..."
emailBodyInfo = ""

'Quick scan command line call - using a user configured profile for the directory and other settings
'WshShell.Run """C:\Program Files\EMC Captiva\QuickScan\quickscn.exe""" &" /scan profile=TEAMS_Import - FileImport", 2, false

Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
'Steps to take to automate

Do While True
	'Check all folders exist
	If Not objFSH.FolderExists(logDirectory) Then
		emailBodyInfo = logDirectory & " folder can not be found." & vbCrLf
	End If
	If Not objFSH.FolderExists(toBeProcessDirectory) Then
		emailBodyInfo = toBeProcessDirectory & " folder can not be found." & vbCrLf
	End If
	If Not objFSH.FolderExists(processingDirectory) Then
		emailBodyInfo = emailBodyInfo & processingDirectory & " folder can not be found." & vbCrLf
	End If
	If Not objFSH.FolderExists(readyToUploadDirectory) Then
		emailBodyInfo = emailBodyInfo & readyToUploadDirectory & " folder can not be found." & vbCrLf
	End If
	If Not objFSH.FolderExists(archiveDir) Then
		emailBodyInfo = emailBodyInfo & archiveDir & " folder can not be found." & vbCrLf
	End If
	If Not objFSH.FolderExists(toBeProcessDirectorySweden) Then
		emailBodyInfo = emailBodyInfo & toBeProcessDirectorySweden & " folder can not be found." & vbCrLf
	End If
	If Not objFSH.FolderExists(readyToUploadDirectorySweden) Then
		emailBodyInfo = emailBodyInfo & readyToUploadDirectorySweden & " folder can not be found." & vbCrLf
	End If

	'Send a mail if any folders are missing
	If emailBodyInfo <> "" Then
		sendEmail emailFrom, emailTo, "There are missing folders - no invoices are being processed", emailBodyInfo
		WScript.Quit(1)
	End If

	'London process
	processFiles filesToProcessList,quickScanArgments, "London"
	'Sweden process
	processFiles filesToProcessListSweden,quickScanArgmentsSweden, "Sweden"

	'Sleep and rotate for new uploads again.
	WScript.Sleep(30000)
Loop

Sub processFiles(colfilesToProcess, quickScanArgs, region)
	KillProcess	processName

	If colfilesToProcess.Count > 0 Then
		LogEvent "Found " & colfilesToProcess.Count & " files to process"
	End If

	'Loop through each file in directory
	For Each file In colfilesToProcess

		LogEvent "Processing File: " & file
		'Move file
		LogEvent "Moving file to: " & processingDirectory
		objFSH.MoveFile file, processingDirectory

		For Each unpfile In unProcessedFileList
			On Error Resume Next
			objFSH.CopyFile unpfile, archiveDir & "\" & region & "_" & StrCleaner(Date, "-") & "_" & StrCleaner(Time, "") & ".pdf"
			LogEvent "Copied: " & unpfile & " to archive for backup renamed to: " & region & "_" & StrCleaner(Date, "-") & "_" & StrCleaner(Time, "") & ".pdf"
		Next

		'Call QuickScan to process
		WshShell.Run quickCmd & quickScanArgs
		LogEvent "Calling quickScan Pro command line: '" & quickCmd & quickScanArgs & "'"

		'Monitor QuickScan process
		ProcessMonitorFinished(processName)

		If currentChecks <= checksBeforeMail Then 'then it was a success.
			LogEvent "Finished processing invoice"
		Else
			KillProcess processName
			WScript.Sleep(3000)
			'Delete any files in the processing folder as it failed
			'we don't want to process the file again.
			LogEvent "Prepare deletion of unprocessed file(s)"
			LogEvent "Executing cmd: " & "del /s/q/f " & processingDirectory & "*.*"
			WshShell.Run "cmd.exe /c del /s/q/f " & processingDirectory & "*.*"
			LogEvent "Finished processing invoice with some errors"
		End If

		'Reset all variables before processing the next file and
		'Process the next file in the toBeProcessed folder
		fileRename = ""
		currentChecks = 0
	Next
End Sub

'**************************METHODS CALLED FROM ABOVE THIS LINE**********************************
Function ProcessMonitorFinished(pProcess)
	On Error Resume Next
	LogEvent "QuickScan Process executed..."
	Do While CheckProcess(pProcess)
		currentChecks = currentChecks + 1
		WScript.Sleep(timeToSleep)

		If currentChecks > checksBeforeMail Then
			LogEvent "QuickScan failed as it has ran for: " & (currentChecks * timeToSleep) / 1000 & " seconds"
			Exit Do
		End if
	Loop
	LogEvent "QuickScan ran for: " & (currentChecks * timeToSleep) / 1000 & " seconds"
	ProcessMonitorFinished = True
End Function

Function CheckProcess(pProcess)
	Dim colProcesses
	Set colProcesses = objWMIService.ExecQuery _
	("Select * from Win32_Process Where Name = '" & pProcess & "'")

	If colProcesses.Count = 0 Then
	  'process is not running.
	  CheckProcess = False
	Else
	  'process is running.
	  CheckProcess = True
	End If
End Function

Sub KillProcess(pProcessName)
	Dim colProcessList
	Set colProcessList = objWMIService.ExecQuery _
		("SELECT * FROM Win32_Process WHERE Name = '" & pProcessName & "'")

	For Each objProcess in colProcessList
		objProcess.Terminate()
	Next
End Sub

Function StrCleaner(strToTidy, charReplaceWith)
Dim objRegExp, outputStr
Set objRegExp = New Regexp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "[(?*"",\\<>&#~%{}+_.@:\/!;]+"
outputStr = objRegExp.Replace(strToTidy, "-")

objRegExp.Pattern = "\-+"
outputStr = objRegExp.Replace(outputStr, charReplaceWith)

StrCleaner = outputStr
End Function

Sub SendEmail(msgFrom, msgTo, msgsubject, msgBody)
	Dim objEMail
	Set objEmail = CreateObject("CDO.Message")
	objEmail.From = msgFrom
	objEmail.To = msgTo
	objEmail.Subject = msgSubject
	objEmail.Textbody = msgBody
	objEmail.Configuration.Fields.Item _
    	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objEmail.Configuration.Fields.Item _
	    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtp
	objEmail.Configuration.Fields.Item _
	    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objEmail.Configuration.Fields.Update
	objEmail.Send
End Sub

Sub LogEvent(message)
    On Error Resume Next

    'If logFileWrite = IsNull Then
    '	sendEmail emailFrom, emailTo, "Error creating log file", "Error creating log file"
    '	'WScript.Quit
    'End If

    If Err = 0 Then
         logFileWriter.WriteLine Date & " " & Time & " - Info: " & message
		 WSCript.Echo Date & " " & Time & " - Info: " & message
    Else
         logFileWriter.WriteLine Date & " " & Time & " - Failure: " & message & vbCrLf & vbCrLf &_
         "Error Description: " & Err.Description & vbCrLf & "Error Code: " & Err.Number
		 WSCript.Echo Date & " " & Time & " - Failure: " & message & vbCrLf & vbCrLf &_
         "Error Description: " & Err.Description & vbCrLf & "Error Code: " & Err.Number
    End If

	Err.Clear
	'logFileWriter.Close
End Sub

'*******************************************************************************
'* File:           ALL DATA IN ONE
'*
'* Purpose:        COPY MULTIPLE CSV FILE CONTENTS INTO ONE CSV FILE FROM TEMPLATE
'* Version:        1.0 (07 February 2013)
'*
'* Author:		   	 Jon Normington
'* Requirements:   > Windows XP
'*                 Windows Script Host 5.6 - CSCRIPT.EXE_OR_WSCRIPT.EXE
'*******************************************************************************

Option Explicit

Dim completedMeasureTemplate, bulkOutput
'----------------------------------------------------------------------------
'			USER CONFIGURATIONS HERE
'----------------------------------------------------------------------------
'Default answers
completedMeasureTemplate = "<CSVTemplateFile>"
dirToProcessFiles = "<DirectoryOfFilesToProcess>"
'----------------------------------------------------------------------------
'			USER CONFIGURATIONS END HERE
'----------------------------------------------------------------------------
Dim objWorkbookBulk, objCopyToSheet, dirToProcessFiles
'Variables for object pointers
Dim objFSO, objShell, objExcel, file, bulkRow
'CSV numeric code
Const XL_CSV = 6
Const BACK_SLASH = "\"
bulkRow = 1

'Objects
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")
Set objExcel = CreateObject("Excel.Application")


'----------------------------------------------------------------------------
'			SCRIPT BEGINS HERE
'----------------------------------------------------------------------------

completedMeasureTemplate = InputBox("Please paste the file path pf the completed measure template file.", "File required to continue", completedMeasureTemplate)
	If completedMeasureTemplate = "" Or Not objFSO.FileExists(completedMeasureTemplate) Then
		objShell.Popup("Completed measure template file doesn't exist or you didn't type anything! Quitting... Sorry")
		WScript.Quit
	End If

dirToProcessFiles = InputBox("Please paste the top-level folder of where your excel files are located to put into one file." & vbCrLf & vbCrLf & "Unfortunately this version can't look into sub-folders please use Windows search to extract all the files into one folder.", "Folder required to proceed", dirToProcessFiles)
'Directory doesn't exist or user put nothing.
If dirToProcessFiles = "" Or Not objFSO.FolderExists(dirToProcessFiles) Then
	objShell.Popup("Folder doesn't exist or you didn't type anything! Quitting... Sorry")
	WScript.Quit
End If

dirToProcessFiles = Trim(dirToProcessFiles)
If StrComp(Mid(dirToProcessFiles, Len(dirToProcessFiles),1), BACK_SLASH) Then
	'Add backslash not to cause any issues
	dirToProcessFiles = dirToProcessFiles & BACK_SLASH
End If

bulkOutput = completedMeasureTemplate & "bulkData.csv"
objShell.Popup("Output of the bulk file will exist at: " & vbCrLf & bulkOutput)

'Save As copy of Template.csv
SaveToCSV completedMeasureTemplate, bulkOutput

'Setup the file to open and input to
objExcel.DisplayAlerts = FALSE
objExcel.Visible = TRUE
Set objWorkbookBulk = objExcel.Workbooks.Open(bulkOutput)
Set objCopyToSheet = objWorkbookBulk.Worksheets(1)

'Loop through all the files in the directory
For Each file In objFSO.GetFolder(dirToProcessFiles).Files
	If UCase(objFSO.GetExtensionName(file.name)) = "CSV" Then
		WScript.Echo "*** Processing *** :" & file.Name
		CopyContents dirToProcessFiles & file.Name
	End If
Next

objWorkbookBulk.Save
objWorkbookBulk.Close

'Copy contents of file processing from row2 to rowX into the new file
Sub CopyContents(fileNameToCopyFrom)
	WScript.Echo "*** CopyContents *** of: " & fileNameToCopyFrom
	Dim objIExcel, objIWorkbookCopyFrom, objCopyFromWorkSheet, r, c
	Set objIExcel = CreateObject("Excel.Application")
	objIExcel.DisplayAlerts = FALSE
	objIExcel.Visible = TRUE

	Set objIWorkbookCopyFrom = objIExcel.Workbooks.Open(fileNameToCopyFrom)
	Set objCopyFromWorkSheet = objIWorkbookCopyFrom.Worksheets(1)

	For r = 2 To objIWorkbookCopyFrom.Worksheets(1).usedrange.rows.count
		bulkRow = bulkRow + 1
		For c = 1 To objIWorkbookCopyFrom.Worksheets(1).usedrange.columns.count
			objWorkbookBulk.Worksheets(1).Cells(bulkRow,c).value  = objIWorkbookCopyFrom.Worksheets(1).Cells(r,c).value
			objWorkbookBulk.Worksheets(1).Cells(bulkRow,16).NumberFormat = "0"
		Next

	Next
	'Close file
	objIWorkbookCopyFrom.Close
	objIExcel.Quit
End Sub


Sub SaveToCSV(fileToOpen, newDirFileName)
	Dim objIExcel, objIWorkbook, objIWorksheet
	Set objIExcel = CreateObject("Excel.Application")
	Set objIWorkbook = objIExcel.Workbooks.Open(fileToOpen)
	objIExcel.DisplayAlerts = FALSE
	objIExcel.Visible = TRUE

	Set objIWorksheet = objIWorkbook.Worksheets(1)
	objIWorksheet.SaveAs newDirFileName, XL_CSV
	objIExcel.Quit
End Sub


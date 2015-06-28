'*******************************************************************************
'* File:           UPDATE XL
'*
'* Purpose:        COMPARE MULTIPLE SHEET/COLUMNS FOR A MATCHING VALUE AND UPDATING THE SHEET
'* Version:        1.0 (05 March 2013)
'*
'* Author:		  	 Jon Normington
'* Requirements:   > Windows XP
'*                 Windows Script Host 5.6 - CSCRIPT.EXE_OR_WSCRIPT.EXE
'*******************************************************************************

Option Explicit

Const fileLoc = "C:\Documents and Settings\user\Desktop\ECA Test 2.xlsx"


Dim fileToProcess, workSheetToCompare, workSheetToUpdate
Dim objFSO, objShell


Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")


fileToProcess = InputBox("Please input your spreadsheet to be processed.", "File required please", fileLoc)
If Len(fileToProcess) = 0 Or Not objFSO.FileExists(fileToProcess) Then
	objShell.Popup("File doesn't exist or you didn't type anything! Quitting... Sorry")
	WScript.Quit
End If

workSheetToUpdate = InputBox("Please input the worksheet to be updated in order as displayed eg. 2", "Sheet to update")
If Len(workSheetToUpdate) = 0 Then
	objShell.Popup("You didn't type anything! Quitting... Sorry")
	WScript.Quit
End If

workSheetToCompare = InputBox("Please input the worksheet number to be compared in order as displayed eg. 3", "Sheet to compare")
If Len(workSheetToUpdate) = 0 Then
	objShell.Popup("You didn't type anything! Quitting... Sorry")
	WScript.Quit
End If


ProcessFile fileToProcess,workSheetToCompare,workSheetToUpdate
WScript.Echo "Finished processing file..."

Sub ProcessFile(file, sheetToCompare, sheetToUpdate)
	Dim objIExcel, objIWorkbook,objWorkSheet, r, i, objWorkSheetToUpdate, objWorkSheetToCompare
	Set objIExcel = CreateObject("Excel.Application")
	objIExcel.DisplayAlerts = FALSE
	objIExcel.Visible = TRUE

	sheetToUpdate = CInt(sheetToUpdate)
	sheetToCompare = CInt(sheetToCompare)

	Set objIWorkbook = objIExcel.Workbooks.Open(file)

'	For i = 1 To objIWorkbook.Worksheets.Count
'		WScript.Echo objIWorkbook.Worksheets(i).Name
'	Next

	Set objWorkSheetToUpdate = objIWorkbook.Worksheets(sheetToUpdate)
	Set objWorkSheetToCompare = objIWorkbook.Worksheets(sheetToCompare)

	Dim prodName, modlNo
	For r = 2 To objIWorkbook.Worksheets(sheetToUpdate).usedrange.rows.count
		prodName = objIWorkbook.Worksheets(sheetToUpdate).Cells(r,2).value
		modlNo = objIWorkbook.Worksheets(sheetToUpdate).Cells(r,3).value
		WScript.Echo prodName & "~" & modlNo

		'Now search sheet x compare to see if there is a match
		For i = 2 To objIWorkbook.Worksheets(sheetToCompare).usedrange.rows.count
			If StrComp(objIWorkbook.Worksheets(sheetToCompare).Cells(i,2).value, prodName) = 0 And _
				StrComp(objIWorkbook.Worksheets(sheetToCompare).Cells(i,3).value, modlNo) = 0 And _
					StrComp(objIWorkbook.Worksheets(sheetToCompare).Cells(i,1).value, "Yes") = 0 Then
				'Update color and input word Yes.
				objIWorkbook.Worksheets(sheetToUpdate).Cells(r,1).value = "Yes"
				objIWorkbook.Worksheets(sheetToUpdate).Cells(r,1).Interior.ColorIndex = 6
				WScript.Echo "Found complete match and updated"
			End If

		Next
	Next
	'Save the new file
	objIWorkbook.Save
	objIWorkbook.Close
	objIExcel.Quit
End Sub

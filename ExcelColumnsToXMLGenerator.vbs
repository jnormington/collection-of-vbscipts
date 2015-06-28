'*******************************************************************************
'* File:           EXCEL COLUMNS TO XML GENERATOR
'*
'* Purpose:        HELP AUTOMATE CREATION OF XML FROM EXCEL FOR A FEED
'* Version:        1.0 (30 Sept 2012)
'*
'* Author:		   	 Jon Normington
'* Requirements:   > Windows XP
'*                 Windows Script Host 5.6 - CSCRIPT.EXE_OR_WSCRIPT.EXE
'*******************************************************************************
Option Explicit
On Error Resume Next

Dim objFSO, objShell, objExcel, objWorkbook, objXMLFile, objWS, objADOStream
Dim xlsFileName, xlsWorkSheet, intFileAttempts, tmpFileName, newLine, intSheetIdx, strWorkSheets, intUserSelection, r, xmlContent, xmlFileName
Const fileNotFoundError = "File to process is not not found. "
xmlFileName = "output.XML"
Const fileExt = ".png"

newLine = Chr(13) & Chr(10)
xlsFileName = "C:\Users\user\Desktop\MD.xls"
intUserSelection = 0
intFileAttempts = 0
r = 1

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")
Set objADOStream = CreateObject( "ADODB.Stream" )
objADOStream.Charset = "utf-8"

'Check if the file exists
If Not objFSO.FileExists(xlsFileName) Then
	objShell.Popup fileNotFoundError & newLine & newLine & xlsFileName &_
					newLine & newLine & " Please input the path to a valid file.", 3
	Do
		'Attempt count for later use in inputbox message
		intFileAttempts = intFileAttempts + 1
		'Depending on the attempt we should display the reason in the inputbox on second attempt.
		If intFileAttempts = 1 Then
			tmpFileName = InputBox("Please paste the path of your excel file", "File required to proceed", xlsFileName)
		Else
			tmpFileName = InputBox("Please paste the path of your excel file." & Chr(13) & Chr(10) &_
								 "The previous file path was not found." &_
									newLine & newLine & tmpFileName, "Valid file required to proceed", tmpFileName)
		End If

		'If the input is "" then just quit as user could have pressed cancel
		If tmpFileName = "" Then
			objShell.Popup "Program exiting due to no file path input.", 3
			QuitCleanly
		End If
	Loop While Not objFSO.FileExists(tmpFileName)

	xlsFileName = tmpFileName
End If

'Set the output file path
xmlFileName = Mid(objFSO.GetAbsolutePathName(xlsFileName), 1, Len(objFSO.GetAbsolutePathName(xlsFileName)) _
										- Len(objFSO.GetFileName(xlsFileName))) & xmlFileName
objShell.Popup "XML OUTPUT: " & xmlFileName, 3

'Ensure no file exists in this place.
If objFSO.FileExists(xmlFileName) Then
	objShell.Popup "Renaming existing file to write to: " & newLine & xmlFileName, 3
	objFSO.GetFile(xmlFileName).Name = "CHARACTERS_" &  Replace(Date, "/", "-") & "_" & Replace(Time, ":", "") & ".XML"
End If

Set objExcel = WScript.CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(xlsFileName,2,False)

'Loop through all the worksheets for display to the user..
For intSheetIdx = 1 To objWorkbook.Worksheets.Count
	Set objWS = objWorkbook.Worksheets.Item(intSheetIdx)
	strWorkSheets = strWorkSheets & " (" & intSheetIdx & ")	" & objWS.Name & newLine
Next
	'Display the worksheets and also the index number to choose
	intUserSelection = WorkSheetSelectionValidation

'Start building text file as XML
objADOStream.Open
objADOStream.WriteText "<?xml version=""1.0"" encoding=""UTF-8""?>" & newLine & "<Characters>"

'Ensure to cast intUserSelection otherwise WorkSheets(Integer) complains
intUserSelection = CInt(intUserSelection)

For r = 2 To objWorkbook.Worksheets(intUserSelection).usedrange.rows.count
	On Error Resume Next
	xmlContent = xmlContent & newLine & "	<Character>"
	If Not objWorkbook.Worksheets(intUserSelection).UsedRange.Columns.Count >= 8 And _
		 objWorkbook.Worksheets(intUserSelection).Cells(r,1) = "" Or _
			objWorkbook.Worksheets(intUserSelection).Cells(r,2) = "" Or _
			objWorkbook.Worksheets(intUserSelection).Cells(r,3) = "" Or _
			objWorkbook.Worksheets(intUserSelection).Cells(r,4) = "" Or _
			objWorkbook.Worksheets(intUserSelection).Cells(r,5) = "" Or _
			objWorkbook.Worksheets(intUserSelection).Cells(r,7) = "" Or _
			objWorkbook.Worksheets(intUserSelection).Cells(r,8) = "" Then
			WScript.Echo "Excel file is not complete ensure all 8 columns from column A-H except LOGO_URL are all populated."
			QuitCleanly
	End If
	xmlContent = xmlContent & newLine & "		<index>" & r - 1 & "</index>"
	xmlContent = xmlContent & newLine & "		<genre-group>" & objWorkbook.Worksheets(intUserSelection).Cells(r,1) & "</genre-group>"
	xmlContent = xmlContent & newLine & "		<header-title>" & objWorkbook.Worksheets(intUserSelection).Cells(r,2) & "</header-title>"
	xmlContent = xmlContent & newLine & "		<sub-header-title>" & objWorkbook.Worksheets(intUserSelection).Cells(r,4) & "</sub-header-title>"
	xmlContent = xmlContent & newLine & "		<thumbnail-displayName><![CDATA[" & objWorkbook.Worksheets(intUserSelection).Cells(r,3) & "]]></thumbnail-displayName>"
	xmlContent = xmlContent & newLine & "		<thumbnail-url>" & objWorkbook.Worksheets(intUserSelection).Cells(r,7) & objWorkbook.Worksheets(intUserSelection).Cells(r,3) & fileExt & "</thumbnail-url>"
	xmlContent = xmlContent & newLine & "		<image-displayName><![CDATA[" & objWorkbook.Worksheets(intUserSelection).Cells(r,3) & "]]></image-displayName>"
	xmlContent = xmlContent & newLine & "		<image-url>" & objWorkbook.Worksheets(intUserSelection).Cells(r,8) & objWorkbook.Worksheets(intUserSelection).Cells(r,3) & fileExt & "</image-url>"
	xmlContent = xmlContent & newLine & "		<logo-url>" & objWorkbook.Worksheets(intUserSelection).Cells(r,6) & "</logo-url>"
	xmlContent = xmlContent & newLine & "		<character-desc><![CDATA[" & objWorkbook.Worksheets(intUserSelection).Cells(r,5) & "]]></character-desc>"
	xmlContent = xmlContent & newLine & "	</Character>"
Next

objADOStream.WriteText xmlContent & newLine & "</Characters>"
objADOStream.SaveToFile xmlFileName
'Finished building XML close file and quit
objShell.Popup "Finising processing excel please check file output is correct.", 5
QuitCleanly

Function WorkSheetSelectionValidation
	On Error Resume Next
	Dim tmpVal, valid

	Do
		tmpVal = InputBox("Please type a worksheet number to process." & newLine &_
										strWorkSheets, "Requires a worksheet index to process")

		If tmpVal = "" Or Len(tmpVal) = 0 Then
			objShell.Popup "No worksheet index selected.. Quitting", 3
			QuitCleanly
		End If

		'Try to cast to Integer
		CInt(tmpVal)
		'Catch the error when attempt to cast
		If Err.Number <> 458 Then
			valid = -1
		Else
			'Valid conditions
			If (CInt(tmpVal) > 0 And CInt(tmpVal) <= objWorkbook.Worksheets.Count) = -1 Then
				valid = 0
			Else
				valid = -1
			End If
		End If
	Loop While valid = -1

	WorkSheetSelectionValidation = tmpVal
End Function

Sub QuitCleanly
	On Error Resume Next
	objWorkbook.Close
	objXMLFile.Close
	objADOStream.Close
	objWorkbook = Nothing
	objXMLFile = Nothing

	WScript.Quit 0
End Sub

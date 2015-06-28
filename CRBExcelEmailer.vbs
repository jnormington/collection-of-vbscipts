'*******************************************************************************
'* File:           CRB EXCEL EMAILER
'*
'* Purpose:        AUTOMATE SENDING EMAILS FROM LDAP LOOKUP FOR INTERNAL CRB CHECKS TO MANAGERS
'* Version:        1.0 (24 May 2012)
'*
'* Author:		   	 Jon Normington
'* Requirements:   > Windows XP
'*                 Windows Script Host 5.6 - CSCRIPT.EXE_OR_WSCRIPT.EXE
'*******************************************************************************
Option Explicit
On Error Resume Next

Const excelFileLocation = "<LocationOfFile>"
Const emailAttachment = "<LocationOfFile>"
Const fileNotFoundError = "File not found to process - please check the file exists"

Const crbSubject = "CRB Reminder"
Const crbBodyHEader = "Dear "
Const crbBody = "The CRB for "
Const crbBodyDate = " is due to expire on "
Const crbBodyFooter = "Please ensure that your member of staff books an appointment at the earliest convienience."
Const crbSignature = "Kind Regards CRB Committee"
Const emailMsgFrom = "<fromEmail>"
Const emailBCC = ""

Dim objFSO, objExcel, objWorkbook, r, emailError, generalErrors, errorMsg
Dim lastName, firstName, crbStartDate, crbEndDate, svLastName, svFirstName, svMailAddress, emMailAddress
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

If Not objFSO.FileExists(excelFileLocation) Then
	WScript.Echo fileNotFoundError
	WScript.Quit
End If

Set objExcel = WScript.CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(excelFileLocation,2,False)

processErrors
For r = 2 To objWorkbook.Worksheets(1).usedrange.rows.count
	lastName = objWorkbook.Worksheets(1).Cells(r,2).value
	firstName = objWorkbook.Worksheets(1).Cells(r,3).value
	crbStartDate = objWorkbook.Worksheets(1).Cells(r,12).value
	crbEndDate = objWorkbook.Worksheets(1).Cells(r,13).value
	svLastName = objWorkbook.Worksheets(1).Cells(r,16).value
	svFirstName = objWorkbook.Worksheets(1).Cells(r,17).value
	draftEmailToManager
	processErrors
Next

objWorkbook.Save
objWorkbook.Close

WScript.Echo "Finished processing all records"
If generalErrors = True Then
	WScript.Echo "Errors Occurred during processing.. " & vbCrLf & errorMsg
End If

Sub draftEmailToManager()
	If emailError = False Then
		emMailAddress = getEmailAddress(firstName, lastName) 'employee email
		svMailAddress = getEmailAddress(svFirstName, svLastName) 'supervisor

		If Not svMailAddress = "" Then
			sendEmail emailMsgFrom, svMailAddress & "," & emMailAddress, emailBCC, crbSubject, crbBody & firstName & " " & lastName & crbBodyDate & crbEndDate &_
					chr(13) & chr(13) & crbBodyFooter & chr(13) & chr(13) & crbSignature, emailAttachment
		End If
	End If

	lastName = ""
	firstName = ""
	crbStartDate = ""
	crbEndDate = ""
	svLastName = ""
	svFirstName = ""
	emailError = False

	processErrors
End Sub

Function getEmailAddress(firstName, lastName)
	Dim objConnection, objCommand, objRecordSet, mailToReturn
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	Err.Clear
	'SQL atatement to query AD replace strNTUserName with
	'the user that is executing this script upon logon
	objCommand.CommandText = _
	    "SELECT givenName, sn, mail " &_
	    "FROM 'LDAP://dc=<ldapHost>,dc=gov,dc=uk' " &_
	    "WHERE objectCategory='user' " &_
	    "AND sn = '" & lastName & "'" &_
	    "AND givenName = '" & firstName & "'"

	Set objRecordSet = objCommand.Execute

	If objRecordSet.BOF Or objRecordSet.EOF Then
		handleDuplicateNamesAndNoRecords False, firstName, lastName
		mailToReturn = ""
	Else If objRecordSet.RecordCount > 1 Then
		handleDuplicateNamesAndNoRecords True, firstName, lastName
		mailToReturn = ""
	Else
		mailToReturn = objRecordSet.Fields("mail")
	End If
	End If

	objConnection.Close
	getEmailAddress = mailToReturn
End Function

Sub handleDuplicateNamesAndNoRecords(duplicate, firstName, lastName)
	emailError = True
	objWorkbook.Worksheets(1).Cells(r,1).Interior.ColorIndex = 3
	objWorkbook.Worksheets(1).Cells(r,25).Interior.ColorIndex = 3

	If duplicate = False Then
		objWorkbook.Worksheets(1).Cells(r,25).value = "NO EMAIL FOUND " & firstName & " " & lastName
	Else
		objWorkbook.Worksheets(1).Cells(r,25).value = "DUPLICATE ENTRY " & firstName & " " & lastName
	End If
End Sub


Sub sendEmail(msgFrom, msgTo, msgBCC, msgsubject, msgBody, fileURL)
	Dim objEmail
	Set objEmail = CreateObject("CDO.Message")
	objEmail.From = msgFrom
	objEmail.To = msgTo

	If Not msgBCC = "" Then
		objEmail.BCC = msgBCC
	End If

	objEmail.Subject = msgSubject
	objEmail.TextBody = msgBody

	If Not fileURL = "" Then
		objEmail.AddAttachment fileURL
	End If

	objEmail.Configuration.Fields.Item _
    	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objEmail.Configuration.Fields.Item _
	    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "<smtpServer>"
	objEmail.Configuration.Fields.Item _
	    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objEmail.Configuration.Fields.Update
	objEmail.Send

	'objEmail = Nothing
End Sub

Sub processErrors
	If Err <> 0 Then
		generalErrors = True
		errorMsg = errorMsg & vbCrLf & Err.Description
	End If
	Err.Clear
End Sub

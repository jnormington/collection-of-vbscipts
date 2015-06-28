'*******************************************************************************
'* File:           UNIVERSAL LOGON SCRIPT
'*
'* Purpose:        TO MAP THE CORRECT DFS SHARE THAT X HAS
'* 				   			 SET UP BASED ON THE USERS AD DEPARTMENT
'*
'* Version:        1.1.0 (01 April 2011)
'*
'* Author:		   	 Jon Normington
'*
'* Version Log
'*		1.0.0 - 01/April/2011 - First release of the script remove S,Y,T drives and remapping and
'*								looking at AD Department to map the T drive back based on collection
'*		1.1.0 - 13/May/2011 - Add ability to retrieve all groups a user is a memberOf and created a function to pass
'*							  the group that you want to check for to provide flexiblility if more special cases like
'*							  Zippy arise in the future for easy update. Also added new LONDON_ZIPPY unc path
'*							  and a new collection for the support of the custom group checks. Also we now close the
'*							  connection to the database
'*
'*
'*******************************************************************************

'If Errors occurs during runtime continue without throwing an exeception
On Error Resume Next

'Department and Special UNC paths
Const MAPPING_CONST = "\\server\dfsshare\london\1"
Const MAPPING_CONST2 = "\\server\dfsshare\london\2"


'The above departments in a collection
Dim intDepartmentsCount
Dim colDepartments(6,2)
intDepartmentsCount = 5

colDepartments(0,0) = "GROUP NAME IN AD"
colDepartments(0,1) = MAPPING_CONST
colDepartments(1,0) = "GROUP NAME IN AD2"
colDepartments(1,1) = MAPPING_CONST2

'*************************************************
'Special Cases for other potential drive mappings
'*************************************************
Dim customGroups(4)
customGroups(0) = "z1"
customGroups(1) = "z2"
customGroups(2) = "z3"
customGroups(3) = "z4"


'______________________________________________________________
'Don't modify anything below this line - The script starts here

Dim strNTUserName
Dim strDepartment
Dim intIndex
Dim colGroups()

intGroupsCount = 0

'Variables holding objects
Dim objWshShell
Dim objWshNetwork
Dim objConnection
Dim objCommand
Dim objValue
Dim objRecordSet

Set objWshShell = CreateObject("WScript.Shell")
LogEvent "Logon Script has started"

'Get the user name of the user
Set objWshNetwork = CreateObject("WScript.Network")
LogEvent "Attempt to retrieve the current logged on user"
strNTUserName = objWSHNetwork.UserName

'Remove the global network mappings
objWshNetwork.RemoveNetworkDrive "S:",True,True
	LogEvent "Attempt to remove the S: drive"
objWshNetwork.RemoveNetworkDrive "Y:",True,True
	LogEvent "Attempt to remove the Y: drive"
objWshNetwork.RemoveNetworkDrive "T:",True,True
	LogEvent "Attempt to remove the T: drive"

objWshNetwork.MapNetworkDrive "S:", LONDON_LONPUB
	LogEvent "Attempt to add S:" & LONDON_LONPUB & "drive"
objWshNetwork.MapNetworkDrive "Y:", LONDON_LONPST
	LogEvent "Attempt to add Y:" & LONDON_LONPST & "drive"

If strNTUserName <> "" Then
	'Make the connection using ADODB object
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	Err.Clear
	'SQL atatement to query AD replace strNTUserName with
	'the user that is executing this script upon logon
	objCommand.CommandText = _
	    "SELECT Department, memberOf " &_
	    "FROM 'LDAP://dc=<ldapHost>,dc=com' " &_
	    "WHERE objectCategory='user' " &_
	    "AND samAccountName = '" & strNTUserName & "'"

	Set objRecordSet = objCommand.Execute
	strDepartment = objRecordSet.Fields("Department")
	colMemberOf = objRecordSet.Fields("memberOf").Value

		If (Err.Number <> 0) Then
			LogEvent "Attempt to retrieve user department from AD"
		Else
			LogEvent "Attempt to retrieve user department from AD" &_
						" and the department is: " & strDepartment
			'Calling method to map drive
			MapTheDepartmentDrive
		End If

		'Checking if the user is part of a Zippy group
		If isUserInGroup(customGroups) Then
			LogEvent "Check if user is part of a Zippy Group: TRUE "
			objWshNetwork.RemoveNetworkDrive "Q:",True,True
			LogEvent "Attempt to remove the Q: drive"
			objWshNetwork.MapNetworkDrive "Q:", LONDON_ZIPPY
			LogEvent "Attempt to add Q:" & LONDON_ZIPPY & "drive"
		Else
			LogEvent "Check if user is part of a Zippy Group: FALSE"
		End If

		'Quit the script when finished
		LogEvent "Logon Script has finished quiting"
		'Close the connection
		objConnection.Close
		WScript.Quit
Else
	LogEvent "Retrieve NTUserName to map department drive"
	WScript.Quit(0)
End If

'_________________________________________________________
' Methods called from the above main - DON'T EDIT

Function isUserInGroup(groupsToCheck)
	Dim bolFoundComma
	Dim memberCheck

	'Set to false so the statement below gets executed
	bolFoundComma = False
	'Loop through each group returned by recordSet and do StrComparision to find the first comma
	'once the first comma is found or if the 4th character is a star then it is a distribution
	'mailbox and we should also ignore.
	For each member in colMemberOf
		'WScript.Echo member
		For i=1 To Len(member)
			If StrComp(Mid(member,i,1),",", 1) = 0 And bolFoundComma = False And Not StrComp(Mid(member,4,1), "*", 1) = 0 Then
				bolFoundComma = True
				For Each groupCheck In groupsToCheck
					'WScript.Echo groupCheck
					If StrComp(groupCheck, Mid(member,4,i - 4), 1) = 0 Then
						isUserInGroup = True
					End If
				Next
			End If
		Next
		'Set back to false for the next group check
		bolFoundComma = False
	Next
End Function

Sub MapTheDepartmentDrive
	On Error Resume Next
	Dim bolMatchFound
	bolMatchFound = False
	For intIndex = 0 To intDepartmentsCount
		If colDepartments(intIndex,0) = strDepartment Then
			objWshNetwork.MapNetworkDrive "T:", colDepartments(intIndex,1)
			LogEvent "Attempt to add T:" & colDepartments(intIndex,1) & "drive"
			bolMatchFound = True
		End If
	Next

	If bolMatchFound <> True Then
		Err.Number = -2000
		LogEvent "Attempt to map the T: drive for deparment: " & strDepartment &_
				" failed as there was no match found." & vbCrLf & vbCrLf &_
				"Please contact helpdesk for further information"
	End If
End Sub

Sub LogEvent(message)
    On Error Resume Next
    If Err = 0 Then
         objWshShell.LogEvent 4, "Logon Script Success: " & message
    Else
         objWshShell.LogEvent 1, "Logon Script Failure: " & message & vbCrLf & vbCrLf &_
         "Error Description: " & Err.Description & vbCrLf & "Error Code: " & Err.Number
    End If
Err.Clear
End Sub


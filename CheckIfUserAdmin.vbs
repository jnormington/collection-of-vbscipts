'*******************************************************************************
'* File:           CHECK IF USER ADMIN - TEMPLATE USED IN OTHER SCRIPTS
'*
'* Purpose:        CHECK IF USER ACCOUNT IS ADMIN OR IN A HELPDESK GROUP
'* Version:        1.0 (02 December 2010)
'*
'* Author:		   	 Jon Normington
'* Requirements:   > Windows XP
'*                 Windows Script Host 5.6 - CSCRIPT.EXE_OR_WSCRIPT.EXE
'*******************************************************************************

Dim bolAdmin
bolAdmin = CheckIfUserAdmin()

'========Check if user is an adminstrator before continuing

Function CheckIfUserAdmin
	Dim bolIsAdmin
	Dim objNetwork
	Dim strComputer
	Dim objGroup
	Dim objUser
	Dim strUser
	Dim strGroup
	Dim authUsers(20)
	Dim upperstrUser
	Dim strHelpdesk
	Dim strTmpUser

	'Add new helpdesk users here
	'this must be uppercased
	authUsers(0) = "JNORMINGTON"

	'Create network object and comp & user is local
	Set objNetwork = CreateObject("Wscript.Network")
	strComputer = objNetwork.ComputerName
	strUser = objNetwork.UserName
	strGroup = "Administrators"

	'Can't check if user is in a group so array holds
	'helpdesk users
	upperstrUser = UCase(strUser)
	For Each strHelpdesk in authUsers
		If strHelpdesk = upperstrUser Then
			bolIsAdmin = true
		End If
	Next


	'Loop through the admin group on the local machine and
	'check if the user logged on in the admin group
	Set objGroup = GetObject("WinNT://" & strComputer & "/" & strGroup)
	For Each objUser in objGroup.Members
		strTmpUser = UCase(objUser.Name)
		If  strTmpUser = upperstrUser Then
			bolIsAdmin = true
		End If
	Next

	If bolIsAdmin = true Then
		CheckIfUserAdmin = "True"
	Else
		CheckIfUserAdmin = "False"
	End If

End Function

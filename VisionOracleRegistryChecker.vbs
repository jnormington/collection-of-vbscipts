'*******************************************************************************
'* File:           Vision and Oracle registry checker
'*
'* Purpose:        Check if Vision III/Oracle 10g/Discoverer 9i is installed.
'* Version:        1.0 (12 November 2010)
'*
'* Author:         Jon Normington
'* Requirements:   > Windows XP
'*                 Windows Script Host 5.6 - CSCRIPT.EXE_OR_WSCRIPT.EXE
'*******************************************************************************

'Check if Oracle Discoverer or Vision III was installed.
'This is a quick tool for the helpdesk to check
'if Vision III or Oracle Discoverer was installed
Option Explicit

Dim wshShell
Dim VisionInst
Dim DiscovInst

Set wshShell = WScript.CreateObject("WScript.Shell")

On Error Resume Next

Const VisionReg = "HKEY_LOCAL_MACHINE\SOFTWARE\ORACLE\KEY_OraClient10g_home1\ORACLE_HOME"
Const DiscovReg = "HKEY_LOCAL_MACHINE\SOFTWARE\ORACLE\HOME0\ORACLE_HOME"


'Check if Oracle Discoverer installed
wshShell.RegRead(DiscovReg)

If Err <> 0 Then
	Err.Clear
	DiscovInst = False
Else
	DiscovInst = True
End If


'Check if Vision III is installed
wshShell.RegRead(VisionReg)

If Err <> 0 Then
	Err.Clear
	VisionInst = False
Else
	VisionInst = True
End If


MsgBox "Is Vision III installed:  " & VisionInst &_
		vbCrLf & "Is Oracle 9i Discoverer installed:  " &_
		DiscovInst, vbInformation ,"Check if Vision III & Oracle Discoverer installed"

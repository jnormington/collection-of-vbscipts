'*******************************************************************************
'* File:      Font Replicator
'*
'* Purpose:   SCRIPT TO REPLICATE FONTS FOR THE RENDER FARM
'* Version:   1.0.0 (21 October 2009)
'*
'* Author:    Jon Normington
'*******************************************************************************

Dim tmpfolder, objfso, strfldr, createfolder, wshlNtwk, strFontFldr, objFolder

tmpfolder = "C:\tmpfontfolder\"
strMaster = "\\share\DeadlineRepository\systm\*"
strFontFldr = "C:\Windows\Fonts"

'*** Checking if tmpfolder for fonts is there if not create ***
Set objfso = WScript.CreateObject("Scripting.FilesystemObject")

If objfso.FolderExists(tmpfolder)=False Then
  WScript.Echo("Folder doesn't exist")
  Createfolder = objfso.CreateFolder(tmpfolder)
  WScript.Echo("Folder Created")
Else
  WScript.Echo("Folder exists")
End If

'*** Download fonts from LONREN15 ***
WScript.Echo("Downloading fonts from master temp location")
objfso.CopyFile strMaster, "C:\tmpfontfolder\", True
WScript.Echo("Fonts copied from LONREN15")

'*** Install the fonts ***
Const FONTS = &H14&
Set objShl = WScript.CreateObject("Shell.Application")
Set objFolder = objShl.NameSpace(FONTS)

objFolder.CopyHere "C:\tmpfontfolder\*.ttf"


' TODO
'   - Delete tmpfont folder after completion
'	  - Compare fonts in the font folder to the fonts in the tmpfont folder
'		and only install the fonts that are not installed - dynamic varible required.
'
'	  - Output the installation of the fonts in event viewer or in a text file...

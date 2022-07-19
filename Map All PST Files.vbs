On error resume next
	
Dim objFSO
Dim objWSH
Dim objWMIService
Dim colProcess
Dim objProcess
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWSH = CreateObject("Wscript.Shell")
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Dim strUserProfile
Dim OutlookOpen
strUserProfile = objWSH.ExpandEnvironmentStrings("%USERPROFILE%")

objWSH.Run "%COMSPEC% /c DIR " & CHR(34) & strUserProfile & "\*.pst" & CHR(34) & " /b /s>" & CHR(34) & strUserProfile & "\My Documents\PST Locations.txt""", 0, True
objWSH.Run "%COMSPEC% /c DIR H:\*.pst /b /s>>" & CHR(34) & strUserProfile & "\My Documents\PST Locations.txt""", 0, True

If Not objFSO.FileExists(strUserProfile & "\My Documents\PST Locations.txt") Then
	Done()
End If
	
Dim AppOutlook, OutlookNS, Path, objfile, Line

Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process")
For Each objProcess In colProcess
	If objProcess.Name = "OUTLOOK.EXE" Then OutlookOpen = True
Next


Set AppOutlook = CreateObject("Outlook.Application")
If err.number <> 0 Then
	Set AppOutlook = GetObject(,"Outlook.Application")
End If

Set OutlookNS = AppOutlook.GetNameSpace("MAPI")
	
Set objfile = objFSO.opentextfile(strUserProfile & "\My Documents\PST Locations.txt")
Do Until objfile.AtEndOfStream
Line=objfile.readline
If instr(Line, "No PST mappings found") Then
	Done()
Else
	OutlookNS.AddStore Line
End If
Loop

If Not OutlookOpen Then
	AppOutlook.Session.Logoff
	AppOutlook.Quit
End If
objfile.close
Set objfile = Nothing
Set AppOutlook = Nothing
Set OutlookNS = Nothing

Done()

Sub Done()
	MSGBOX "Remapping of PST(s) is complete.  Please ensure that your PST(s) have been remapped succesfully."
	If Not OutlookOpen Then objWSH.Run "%COMSPEC% /c Start Outlook",0,False
	objFSO.DeleteFile(strUserProfile & "\Desktop\Map All PST Files.vbs")
	Wscript.Quit
End Sub
@title Profile Data Backup Tool v1.7
@echo off
echo.
goto :beginning

=============================================================================================================
NOTES & UPDATES

v1.0 Back up the profile registry key, network drives and printers
	v1.1 Backup the EFS cert
	v1.2 EFS Cert is promping for a password. Added a line that creates a password within the script so the user doesn't have to do anything
	v1.3 Remove the local printers from the network drives and remove extraneous info from network drives export
	v1.4 Split the Nertwork Drives and Printers into their own seperate files
	v1.5 Export registry keys for the icons on the taskbar and start menu
	v1.6 Added a restart command to prevent delays during the rebuild process. Most of the time it requires a restart anyway.
	v1.7 Backup the Desktop background registry key so the new profile can be even closer to the original.

=============================================================================================================

:beginning
REM backup EFS cert
echo Backing up your File Encryption Certificate
	powershell -command '$secret=""' >nul
	powershell -command "Get-ChildItem cert:\currentuser\my | Where-Object { $_.HasPrivateKey -and $_.PrivateKey.CspKeyContainerInfo.Exportable } | Foreach-Object { [system.IO.file]::WriteAllBytes( '%USERPROFILE%\Desktop\File Encryption Certificate.pfx' , ($_.Export('PFX', $secret )) ) }" >nul

REM back up network drives
echo Backing up your network drive paths.

net use > %USERPROFILE%\Desktop\Networkdrive.txt
for /f "skip=6 delims=*" %%a in (%USERPROFILE%\Desktop\Networkdrive.txt) do ( 

REM deleting the first 6 lines
	echo %%a >>%USERPROFILE%\Desktop\del_Networkdrive.txt )

REM make a new copy of the network drive file after deleting the lines
xcopy %USERPROFILE%\Desktop\del_Networkdrive.txt %USERPROFILE%\Desktop\Networkdrive.txt /y 

REM find the string and delete them and make a new copy of file with just the network drives
findstr /v "OK" %USERPROFILE%\Desktop\del_Networkdrive.txt > %USERPROFILE%\Desktop\Networkdrive.txt 

REM find the string and delete them and make a new copy of file with just the network drives
findstr /v "Unavailable" %USERPROFILE%\Desktop\del_Networkdrive.txt > %USERPROFILE%\Desktop\Networkdrive.txt 

REM find the string and delete them and make a new copy of file with just the network drives
findstr /v "Microsoft Windows Network" %USERPROFILE%\Desktop\del_Networkdrive.txt > %USERPROFILE%\Desktop\Networkdrive.txt

REM find the string and delete them and make a new copy of file with just the network drives
findstr /v "The command completed successfully." %USERPROFILE%\Desktop\del_Networkdrive.txt > %USERPROFILE%\Desktop\Networkdrive.txt

powershell -Command "(gc %USERPROFILE%\Desktop\Networkdrive.txt) -replace 'Microsoft Windows Network', '' | Out-File %USERPROFILE%\Desktop\Networkdrive.txt"
powershell -Command "(gc %USERPROFILE%\Desktop\Networkdrive.txt) | ? {$_.trim() -ne '' } | set-content %USERPROFILE%\Desktop\Networkdrive.txt"

del %USERPROFILE%\Desktop\del_Networkdrive.txt /f /q > nul

REM back up network printers
echo Backing up your printer paths
	powershell -Command "get-WmiObject -class Win32_printer | ft Name > %USERPROFILE%\Desktop\Networkprinters.txt
	powershell -Command "(gc %USERPROFILE%\Desktop\Networkprinters.txt) -replace 'Name', '' | Out-File %USERPROFILE%\Desktop\Networkprinters.txt"
	powershell -Command "(gc %USERPROFILE%\Desktop\Networkprinters.txt) -replace '----', '' | Out-File %USERPROFILE%\Desktop\Networkprinters.txt"
	powershell -Command "(gc %USERPROFILE%\Desktop\Networkprinters.txt) -replace 'Microsoft XPS Document Writer', '' | Out-File %USERPROFILE%\Desktop\Networkprinters.txt"
	powershell -Command "(gc %USERPROFILE%\Desktop\Networkprinters.txt) -replace 'Fax', '' | Out-File %USERPROFILE%\Desktop\Networkprinters.txt"
	powershell -Command "(gc %USERPROFILE%\Desktop\Networkprinters.txt) -replace ' ', '' | Out-File %USERPROFILE%\Desktop\Networkprinters.txt"
	powershell -Command "(gc %USERPROFILE%\Desktop\Networkprinters.txt) -replace 'AdobePDF', '' | Out-File %USERPROFILE%\Desktop\Networkprinters.txt"
	powershell -Command "(gc %USERPROFILE%\Desktop\Networkprinters.txt) -replace 'Snagit12', '' | Out-File %USERPROFILE%\Desktop\Networkprinters.txt"
	powershell -Command "(gc %USERPROFILE%\Desktop\Networkprinters.txt) -replace 'SendToOneNote2010', '' | Out-File %USERPROFILE%\Desktop\Networkprinters.txt"
	powershell -Command "(gc %USERPROFILE%\Desktop\Networkprinters.txt) | ? {$_.trim() -ne '' } | set-content %USERPROFILE%\Desktop\Networkprinters.txt"

REM backup the registry key
echo Backing up your current profile registry key
set THEME_REGKEY="HKLM\SOFTWARE\Wow6432Node\Network Associates\ePolicy Orchestrator\Agent" > nul
set THEME_REGVAL=LoggedOnUser > nul

REM Check for presence of key first.
reg query %THEME_REGKEY% /v %THEME_REGVAL% || (echo No theme name present! & pause & exit /b 1) > nul

REM query the value. pipe it through findstr in order to find the matching line that has the value. only grab token 3 and the remainder of the line. %%b is what we are interested in here.
set THEME_NAME=
for /f "tokens=2,*" %%a in ('reg query %THEME_REGKEY% /v %THEME_REGVAL% ^| findstr %THEME_REGVAL%') do (set THEME_NAME=%%b) > nul

REM replace any spaces with +
set THEME_NAME=%THEME_NAME: =+%
wmic useraccount where (name='%THEME_NAME%') get sid >SID.txt
powershell -Command "(gc SID.txt) -replace 'SID', '' | Out-File SID.txt"
powershell -Command "(gc SID.txt) -replace ' ', '' | Out-File SID.txt"
powershell -Command "(gc SID.txt) | ? {$_.trim() -ne '' } | set-content SID.txt"
for /f "delims=" %%x in (SID.txt) do set SID=%%x
reg export "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\%SID%" "%Userprofile%\contacts\oldprofileBackupKey.reg"
reg export "HKU\%SID%\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband" "%Userprofile%\contacts\taskbarRestore.reg"
reg export "HKU\%SID%\Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage2" "%Userprofile%\contacts\originalStartMenuRestore.reg"
reg export "HKCU\Control Panel\Desktop" "%Userprofile%\contacts\originalDesktopRestore.reg"
%USERPROFILE%\Contacts\backupPopups2.vbs
del %USERPROFILE%\Contacts\backupPopups2.vbs
shutdown.exe /r /f /t 0
(goto) 2>nul & del "%~f0"
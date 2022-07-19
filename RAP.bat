@title Rebuild a Profile Tool v4.31
@echo off
goto :beginning


=============================================================================================================
UPDATES AND NOTES:

v1.0 Created the base of the program that checked the size of the NTUSER.dat file on a remote machine after putting in the machine name and username.
		v1.01 Fixed a bug where the the program would exit if either the machine or user was not found. Now it prompts you again until it finds both.
	v1.1 Created the backup file portion of the program. It's copied to the user's desktop. Backs up the EFS cert, network drives and network printers.
	v1.2 Set a RAPPath variable that ensures the program will function correctly even when a 'cd' command is introduced.
	v1.3 Added text instructions to the program so users other than me who run it will know what to do
	v1.4 Added text instructions 
	v1.5 Added a part at the beginning to establish who the SvD agent logged into the computer is. It kept erroring out because it was looking in the adm account's profile.
	v1.6 Copy the SID file from the user's desktop to the RAP folder for future idea of remote registry key deletion.
	v1.7 Remotely rename the local profile of the user.
	v1.8 I realized that most of the time I would have to restart the computer. If the machine needs to reboot press pause. Inserted a pause command to wait for the machine to reboot. Press enter to continue.
	v1.9 Updated the renaming convention to include the time

v2.0 Added prompts to continue that would only accept y, Y, n, or N as valid inputs
	v2.1 Removed the portion that backed up the profile registry key to SvD computer
	v2.2 Copy over the Profile Migration Tool to the user's desktop
	v2.3 Copy the Map All PSTs tools to the user's desktop.
	v2.4 After a lot of rebuilds, I realized there are instances where the same user has had a name change and the tool no longer runs properly when the new profile name is different. Added a portion that reprompts the agent for the username incase it has changed.
	v2.5 Added a portion that looks for the new user's profile folder and continues automatically after it is detected rather than waiting for the agent to hit a button to continue.
	v2.6 clear the screen after a rebuild is complete
	v2.7 Added a section that copied a log of what was done to the clipboard
	v2.8 Added a vbs message box that alerted the user to paste the clipboard contents into the ticket prior to hitting continue
	v2.9 Launch an MSRA session automatically as the last step

v3.0 Added the functionality to remotely delete the registry from the tool
	v3.1 Removed the time of the rebuild from the local profile name
	v3.2 Added a part that reset the remotePC and remoteUser variables to ones that would never exist anywhere ever...
	v3.3 - 3.6 Modified on-screen instructions to reflect what needed to be done by the agent
	v3.7 Created a part that automatically continues after the machine has restarted. The backup tool now automatically restarts the computer.
	v3.8 Changed the way the NTUSER.dat file size is displayed. It is now displayed in whole integer MB instead of bytes.
	v3.9 Changed the way the NTUSER.dat file size is displayed. It is now displayed in MB rounded to two decimal places to allow a better informed decision about whether or not to continue.

v4.0 The tool now looks at the registry to detect the logged on user. If it can't detect it from the registry, it gives the option for the agent to manually enter it. It will display the contents of the remote computer's Users folder to ensure no misspellings occur. It will loop until a valid username is inputted.
		v4.01 Updated echoed text to minimize screen space
	v4.1 Copy the script to remap drives and printers to the user's contacts folder.
	v4.2 Look for the registry key value of the logged in user after the profile is renamed and set the value of that key as the new remoteUser variable. If the remoteUser variable equals "N/A", the program will loop until it finds a user that is NOT N/A. Once it has done this it will wait until it verifies that the logged in user's profile folder has been created.
		v4.201 Fixed a minor bug that caused the program to crash when it attempted to compare the remoteUser variable to a string. Forgot to put in a second equals sign. DUH!!!
		v4.202 Updated echo text for easier readability
		v4.21  Force the deletion of the remote registry key without a prompt
		v4.211 Added nul commands to make all the commands silent
	v4.3 Remove SYSTEM Flag from User's Profile just in case it is there
		v4.31  Delete local SID text file
=============================================================================================================

:beginning
set THEME_REGKEY="HKLM\SOFTWARE\Wow6432Node\Network Associates\ePolicy Orchestrator\Agent"
set THEME_REGVAL=LoggedOnUser

REM Check for presence of key first.
reg query %THEME_REGKEY% /v %THEME_REGVAL% > nul || (echo User not found! & pause>nul & exit /b 1)

REM query the value. pipe it through findstr in order to find the matching line that has the value. only grab token 3 and the remainder of the line. %%b is what we are interested in here.
set THEME_NAME=
for /f "tokens=2,*" %%a in ('reg query %THEME_REGKEY% /v %THEME_REGVAL% ^| findstr %THEME_REGVAL%') do (
    set THEME_NAME=%%b
) > nul

set RAPPath="\\127.0.0.1\C$\Program Files\HDUTILS\RAP"

:connectIn
cls
set remotepc= SuperRealComputerName
set remoteuser=Iamdef.Inetly-Real
	set /p remotepc= Enter the remote computer name or IP address: 
if exist \\%remotepc%\C$\ goto :getUsername
echo.
echo Check the machine name. If it is correct, either use the IP address or instruct user to restart the machine and log in.
echo.
echo Restarting in 3 seconds...
SET tmout=3
PING 1.2.1.2 -n 1 -w %tmout%000 > NUL
goto :connectIn

:getUsername
set RemoteREGKEY=HKLM\SOFTWARE\Wow6432Node\Network Associates\ePolicy Orchestrator\Agent
set RemoteREGVAL=LoggedOnUser

REM Check for presence of key first.
reg query "\\%remotepc%\%RemoteREGKEY%" /v %RemoteREGVAL% > nul || (echo No one is logged into this machine! & pause>nul & exit /b 1)

REM query the value. pipe it through findstr in order to find the matching line that has the value. only grab token 3 and the remainder of the line. %%b is what we are interested in here.
set THEME_NAME=
for /f "tokens=2,*" %%a in ('reg query "\\%remotepc%\%RemoteREGKEY%" /v %RemoteREGVAL% ^| findstr %RemoteREGVAL%') do (
    set remoteuser=%%b
) >nul
if exist \\%remotepc%\C$\Users\%remoteuser%\ goto :choice
goto :inputUsername

:inputUsername
dir \\%remotepc%\C$\Users\
	set /p remoteuser= Enter remote user's username: 
if exist \\%remotepc%\C$\Users\%remoteuser%\ goto :choice
echo Invalid username. Please verify the spelling and try again.
SET tmout=3
PING 1.2.1.2 -n 1 -w %tmout%000 > NUL
goto :inputUsername
echo.


:choice
echo.

set size=0
call :filesize "\\%remotepc%\C$\Users\%remoteuser%\NTUSER.DAT"

REM set filesize of 1st argument in %size% variable, and return
:filesize
  set size="%~z1/1"
  set /a sizeWhole="%size%/1024/1024"
  set /a sizeDec="%size%/1024"
  set /a sizeDec2=%sizeDec:~-3%
  set /a Fraction ="%sizeDec2%/10"

echo %remoteuser%'s NTUSER.dat size is %sizeWhole%.%Fraction% MB
echo Profile corruption typically occurs around 5 MB. 
echo.
set /P continue=Do you want to continue rebuilding this profile?[Y/N] 
if /I "%continue%" EQU "Y" goto :sendBackup
if /I "%continue%" EQU "y" goto :sendBackup
if /I "%continue%" EQU "N" goto :abandon
if /I "%continue%" EQU "n" goto :abandon
echo.
echo Error: Invalid Parameter. Please press either "y" or "n" to continue.
echo.
goto :choice

:abandon
cls
echo Standby. Restarting program...
SET tmout=3
PING 1.2.1.2 -n 1 -w %tmout%000 > NUL
goto :connectIn

REM Copy over the Backup file

:sendBackup
 echo Copying backup tools to remote user's desktop
 set remoteuserdesktop=\\%remotepc%\C$\Users\%remoteuser%\Desktop
 	robocopy %RAPPath%\ %remoteuserdesktop%\ Backup.bat > nul
 	robocopy %RAPPath%\ \\%remotepc%\C$\Users\%remoteuser%\Contacts backupPopups2.vbs > nul
echo.
echo.
echo.
echo Instruct the user to run the file named Backup.bat on their desktop and 
echo let you know when it finishes. When it is done running, the command window 
echo will close and 3 files should have appeared on the Desktop. Tell them to
echo select NO when prompted for a Green Shutdown and inform you when they see 
echo			the Ctrl + Alt + Del screen but...
echo.
echo			DO NOT LOG BACK IN!!
echo.

SET tmout=3
PING 1.2.1.2 -n 1 -w %tmout%000 > NUL



REM Check if the machine is online yet. The remote computer will restart. This waits until it is back online to continue the script

echo %remotepc% is currently restarting; Monitoring for completed reboot; 
echo			Please Standby...

:machineOnline
if exist \\%remotepc%\C$\Users\%remoteuser%\ goto :machineOnline 
echo %remotepc% is offline; Checking for connection to the network

:machineOffline
if exist \\%remotepc%\C$\Users\%remoteuser%\ goto :backOnline
goto :machineOffline

:backOnline
echo.
echo %remotepc% is back Online Now; Continuing with rebuild...


:copySID
robocopy %remoteuserdesktop%\ %RAPPath%\ SID.txt > nul
del %remoteuserdesktop%\SID.txt
echo.

REM Remove SYSTEM Flag from User's Profile
attrib -s -h \\%remotepc%\c$\users\%remoteuser%

:rename
 Set CURRDATE=%TEMP%\CURRDATE.TMP

 Set CURRTIME=%TEMP%\CURRTIME.TMP 
DATE /T > %CURRDATE%

Set PARSEARG="eol=; tokens=1,2,3,4* delims=/, "
 For /F %PARSEARG% %%i in (%CURRDATE%) Do SET YYYYMMDD=%%j%%k%%l

Echo Renaming profile...
RENAME \\%remotepc%\C$\Users\%remoteuser% old_%remoteuser%_%YYYYMMDD% >nul
set oldusername=old_%remoteuser%_%YYYYMMDD%


:choice2
if exist \\%remotepc%\C$\Users\old_%remoteuser%_%YYYYMMDD% goto :regKeyDelete
goto :rename


:regKeyDelete
REM Delete registry key
cd "C:\Program Files\HDUTILS\RAP\"
	echo Deleting the profile registry key of %remoteUser%
	set /p Build=<SID.txt
	reg delete "\\%remotepc%\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\%Build%" /f > nul
echo.
echo. 
echo Waiting on customer login...

:getNewUsername
set RemoteREGKEY=HKLM\SOFTWARE\Wow6432Node\Network Associates\ePolicy Orchestrator\Agent
set RemoteREGVAL=LoggedOnUser

REM Check for presence of key first.
reg query "\\%remotepc%\%RemoteREGKEY%" /v %RemoteREGVAL% > nul || (echo No one is logged into this machine! & pause>nul & exit /b 1)

REM query the value. pipe it through findstr in order to find the matching line that has the value. only grab token 3 and the remainder of the line. %%b is what we are interested in here.
set THEME_NAME=
for /f "tokens=2,*" %%a in ('reg query "\\%remotepc%\%RemoteREGKEY%" /v %RemoteREGVAL% ^| findstr %RemoteREGVAL%') do (
    set remoteuser=%%b
) >nul

if %remoteUser% == N/A goto :getNewUsername
if exist \\%remotepc%\C$\Users\%remoteUser%\ goto :nextplease

:nextplease
set remoteuserdesktop=\\%remotepc%\C$\Users\%remoteuser%\Desktop

:copymigrationTools
echo Copying migration tools to the user's desktop...
robocopy %RAPPath%\ %remoteuserdesktop%\ "Profile Migration Tool.bat" >nul
robocopy %RAPPath%\ \\%remotepc%\C$\Users\%remoteuser%\Contacts remapPrinters.bat >nul
robocopy %RAPPath%\ %remoteuserdesktop%\ "Map All PST files.vbs" >nul

:finish
echo.


(
echo.
echo NTUSER.dat is %sizeWhole%.%Fraction% MB
echo Profile Rebuilt 
echo Backup up EFS Certificates to desktop. 
echo Network drives and printers backed up to a text file on user's desktop. 
echo Backup of the profile registry key exported the old profile's 'Contacts' folder. 
echo Remote machine restarted
echo Removed SYSTEM Flag from User's Profile. [KM004325W]
echo Appended 'old_' to the beginning of the local profile name and the date of rebuild to the end. 
echo Profile registry key deleted. 
echo Agent instructed the user to log in. 
echo Perm to RA 
echo Data copied from old profile, EFS certificates restored, and network drives and printers remapped.
echo Issue resolved? [Y/N]: ) | clip

%RAPPath%\complete.vbs

echo Launching an MSRA session to the user's machine.

msra /offerra %remotepc%
del %RAPPath%\SID.txt

goto :connectIn

:restart
echo Standby. Restarting program...
SET tmout=3
PING 1.2.1.2 -n 1 -w %tmout%000 > NUL
goto :connectIn
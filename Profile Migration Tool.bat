@echo off
@title Profile Migration Tool v4.3
goto :beginning

=============================================================================================================
NOTES & UPDATES

v1.0 Copy over all data from the old profile
	v1.1  Automatically import the EFS cert
	v1.2  Check if the EFS cert actually exists first, then import it. Skip if it doesn't exist
	v1.3  Launch the Map All PSTs file

v2.0 Split the program in half. Documents and Favorites are moved after the EFS cert is imported
	v2.1  Close outlook prior to launching the PSTs map tool

v3.0 Import the exported registry keys for the Start Menu and Taskbar icons.

v4.0 Ridded the need of a text file to find the old profile name. With the removal of the time from the main rename program, a condition was added to find the profile named old_%USERNAME%_%DATE%
	v4.1  Added a line that launched the Drive and Printer Remap tool
	v4.11 Made the import of the registry keys silent
	v4.2  Import the desktop background registry key
	v4.3  Move GPO update command from here to printer and network drive remap tool so the copy process doesn't get hung up
=============================================================================================================

:beginning
set THEME_REGKEY="HKLM\SOFTWARE\Wow6432Node\Network Associates\ePolicy Orchestrator\Agent"
set THEME_REGVAL=LoggedOnUser
cls
REM Check for presence of key first.
reg query %THEME_REGKEY% /v %THEME_REGVAL% || (echo No theme name present! & pause & exit /b 1)
cls
REM query the value. pipe it through findstr in order to find the matching line that has the value. only grab token 3 and the remainder of the line. %%b is what we are interested in here.
set THEME_NAME=
for /f "tokens=2,*" %%a in ('reg query %THEME_REGKEY% /v %THEME_REGVAL% ^| findstr %THEME_REGVAL%') do (
    set THEME_NAME=%%b
)
cls

:rename
 Set CURRDATE=%TEMP%\CURRDATE.TMP

 Set CURRTIME=%TEMP%\CURRTIME.TMP 
DATE /T > %CURRDATE%

Set PARSEARG="eol=; tokens=1,2,3,4* delims=/, "
 For /F %PARSEARG% %%i in (%CURRDATE%) Do SET YYYYMMDD=%%j%%k%%l

set content=c:\Users\old_%THEME_NAME%_%YYYYMMDD%

if exist "%USERPROFILE%\Desktop\File Encryption Certificate.pfx" "%USERPROFILE%\Desktop\File Encryption Certificate.pfx"
robocopy "%content%\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" "%USERPROFILE%\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" /E /COPY:DAT /IS /XA:S /XF *.ini /XJ /R:2 /W:5
robocopy "%content%\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu" "%USERPROFILE%\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu" /E /COPY:DAT /IS /XA:S /XF *.ini /XJ /R:2 /W:5
robocopy %content%\*.* C:\Users\%THEME_NAME%\Documents\ProfileRootFiles /E /COPY:DAT /IS /XA:SHT /XF *.ini /XD "AppData" /XD "Documents" /XD "Favorites" /XD "Local Settings" /XJ /R:2 /W:5
robocopy %content% C:\Users\%THEME_NAME% /E /COPY:DAT /IS /XA:SHT /XF *.ini /XD "AppData" /XD "Documents" /XD "Favorites" /XD "Local Settings" /XJ /R:2 /W:5
if exist "%USERPROFILE%\Desktop\File Encryption Certificate.pfx" "%USERPROFILE%\Desktop\File Encryption Certificate.pfx"

REM Remap the Printers and Drives
	start %USERPROFILE%\Contacts\remapPrinters.bat


REM Import pinned icons from taskbar and start menu
regedit.exe "%USERPROFILE%\Contacts\taskbarRestore.reg"
regedit.exe "%Userprofile%\contacts\originalStartMenuRestore.reg"
regedit.exe "%Userprofile%\contacts\originalDesktopRestore.reg"

robocopy %content%\Favorites C:\Users\%THEME_NAME%\Favorites /E /COPY:DAT /IS /XA:SHT /XF *.ini /XJ /R:2 /W:5
robocopy C:\Users\%THEME_NAME%\ C:\Users\%THEME_NAME%\Documents\ /MOVE /E /XF *.ini /XA:SH /XJD /XJF /XD "AppData" /XD "Contacts" /XD "Desktop" /XD "Downloads" /XD "Favorites" /XD "Links" /XD "Music" /XD "Documents" /XD "Pictures" /XD "Saved Games" /XD "Searches" /XD "Videos" /XD "mcafee dlp quarantined files"
robocopy "%content%\Favorites\Links for United States" "C:\Users\%THEME_NAME%\Favorites\Websites for United States"
robocopy "%content%\Favorites\Links" "C:\Users\%THEME_NAME%\Favorites\Favorites Bar"
robocopy %content%\Documents C:\Users\%THEME_NAME%\Documents /E /COPY:DAT /IS /XA:SHT /XF *.ini /XJ /R:2 /W:5
taskkill /im outlook.exe
wscript "MAP ALL PST Files.vbs" //e:vbscript
(goto) 2>nul & del "%~f0"
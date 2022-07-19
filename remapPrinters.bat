@echo off
@title Printer and drive remapper v1.2
goto :start

=============================================================================================================
NOTES & UPDATES

v1.0 Import the network drives and printers
	v1.1 Updated script to reflect the two seperate files that represent the drives and printers
	v1.2 Move GPO update command from data copy tool to here so the copy process doesn't get hung up

=============================================================================================================

:start
for /f "tokens=1,2" %%a in (%USERPROFILE%\Desktop\Networkprinters.txt) do (start %%a)

net use * /delete /y
for /f "tokens=1,2,3,4" %%a in (%USERPROFILE%\Desktop\Networkdrive.txt) do (net use %%a %%b /persistent:yes)


REM GPO Update
gpupdate /force

exit
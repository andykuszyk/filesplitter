@echo off
color 17
title Andy's Amazing File Splitter

:::                                                      
:::          ##       ###    ##  ####    ##     ##  ##      
:::         ####      ####   ##  ## ##    ##   ##    ##   #####
:::        ##  ##     ## ##  ##  ##  ##    ## ##         ##   #
:::       ########    ##  ## ##  ##  ##     ###          ####
:::      ##      ##   ##   ####  ## ##      ##         #    ## 
:::     ##        ##  ##    ###  ####       ##          ####
:::  ,                     ,        _,                            ,_
::: /-\       |\/|        /-\       /_        |        |\|        |_?
:::  ________        _        _______
::: /@@@@@@@@| (@)  |@|      |@@@@@@@|
::: |@|_____    _   |@|      |@|__
::: |@@@@@@@|  |@|  |@|      |@@@@|
::: |@|        |@|  |@|      |@|
::: |@|        |@|  |@|___   |@|_____
::: |@|        |@|  |@@@@@|  |@@@@@@@|                        _
:::           ________        ___       _                    |@|
:::          /@@@@@@@@\      |@@@\_    |@|              _____|@|______
:::         /@        @\     |@@_@@\   |@|             |@@@@@@@@@@@@@@|
:::        /@          @     |@| \@@\  |@|         (@)       |@|
:::        |@                |@|  |@@| |@|          _        |@|
:::        |@                |@|_/@@/  |@|         |@|       |@|
:::        \@                |@@@@@/   |@|         |@|       |@|
:::         \@               |@@@/     |@|         |@|       |@|
:::          \@-------.      |@|       |@\_______  |@|       |@|
:::           -@@@@@@@@\     |@|       |@@@@@@@@@| |@|       |@|
:::                    @.    |@|       \@@@@@@@@@| |@|       |@|
:::                    @|    
:::                    @|  _ _   _____   _,_   _       _,_   _____   _   _
:::          _         @.  | |     |      |    |        |      |      \_/
:::          @\_______@/   | |     |      |    |        |      |       |
:::           @@@@@@@@@    \_/     |     _|_   |___|   _|_     |       |
:::                                           
::: 
::: 

:lblMenu
call :lblSplash
echo. +-------------------+
echo. ^| AAFSU ^| Main Menu ^|
echo. +-------------------+
echo.
echo. 1) Split a file into pieces.
echo. 2) Join pieces back into a file.
echo. 3) Exit
echo.

:lblChoices
Set /p MenuChoice=Enter option:  

If "%MenuChoice%"=="1" goto lblSplit
If "%MenuChoice%"=="2" goto lblJoin
If "%MenuChoice%"=="3" exit
if "%MenuChoice%"=="exit" exit
echo.
echo Invalid choice, try again.
echo.
goto lblChoices


:lblSplit
call :lblSplash
echo.
echo. +---------------+
echo. ^| AAFSU ^| Split ^|
echo. +---------------+
:lblSplitMenu
echo.
echo Please enter the file name of the file you'd like to split or type "list" 
echo to list files in this folder. Type "exit" to return to the main menu.
echo.
set /p FileName=Enter file name:  

if "%FileName%"=="list" (
	echo.
	dir /b
	echo.
	echo. File list complete...
	echo.
	goto lblSplitMenu
)
if "%FileName%"=="exit" goto lblMenu
if not exist %FileName% (
	echo %FileName% is not a valid file name. Please try again.
	goto lblSplitMenu
)
echo.
echo Now please enter the size of pieces that you would like the file split into.
echo.
set /p NumberOfKBs=Size in KBs (1024KBs = 1MB):  
echo.
echo Thank you, splitting file...

findstr "^:" "%~sf0" | findstr /i /v ":lbl" | findstr /i /v ":::" >temp.vbs 
cscript //nologo temp.vbs "%FileName%" %NumberOfKBs%
del temp.vbs

echo.
echo Split complete!
echo.

goto lblMenu

:lblJoin
call :lblSplash
echo.
echo. +--------------+
echo. ^| AAFSU ^| Join ^|
echo. +--------------+
:lblJoinMenu
echo.
echo Please enter the file name of the files you'd like to join or type "list" 
echo to list files in this folder. Type "exit" to return to the main menu.
echo.
echo You should type the name of the file up until the numbered extensions. e.g:
echo.
echo For test.001, test.002, test.003, etc., you should type "test"
echo.
set /p FileName=Enter file name:  

if "%FileName%"=="list" (
	echo.
	dir /b
	echo.
	echo File list complete...
	echo.
	goto lblJoinMenu
)
if "%FileName%"=="exit" goto lblMenu
if not exist "%FileName%".001 (
	echo %FileName%.001 does not exist so there are no files to join. Please try again.
	goto lbljoinMenu
)
echo.
echo Thank you, you've asked to join the %FileName% files.
echo.
echo Now please enter the file extension of the original file (zip, xls, etc.):
echo.
set /p Extension=Enter file extension:  
echo.
echo Thank you, joining files...
echo.
copy /b %FileName%.* %FileName%.%Extension%
echo.
echo Join complete!
echo.
goto lblMenu


:lblSplash
cls
for /f "delims=: tokens=*" %%A in ('findstr /b ::: "%~f0"') do @echo(%%A
goto :EOF

:'--- Declare vars ---
:LF = Chr(10)
:Dim oFSO, FullName, Path, Name, Size, TargetDir
:Set oFSO = CreateObject("Scripting.FileSystemObject")
:Dim iFile, oFile, iStream, Data
:Dim Ext, e, offset, length, NewName

:'--- Set vars ---
:FullName=Wscript.Arguments.Item(0)
:Path = oFSO.GetParentFolderName(FullName)
:Name = oFSO.GetBaseName(FullName) & "."
:'Enter size in Kilobytes
:Size = Wscript.Arguments.Item(1) * 1024
:TargetDir=Path
:Ext = 0
:offset = 1

:'--- Set input file as specified file, FullName ---
:Set iFile = oFSO.GetFile(FullName)

:'--- Set input stream object as text stream of input file --- ??? Text stream works ok?
:Set iStream = iFile.OpenAsTextStream(1)

:'--- Set Data as file size? ---
:Data = iStream.Read(iFile.Size)
:iStream.close

:'--- Splitting loop ---
:Do
:	'--- Control file extension ---
:	Ext = Right("00" & Ext + 1, 3)
:	
:	'--- Set name of new output file ---
:	NewName = TargetDir & Name & Ext
:	
:	'--- Set output file ---
:	Set oFile = oFSO.CreateTextFile(NewName, 2)
:
:
:	length = Size
:	If length > Len(data)+1 - offset Then length = Len(data) + 1 - offset
:
:	oFile.Write Mid(Data, offset, length)
:	offset = offset + length
:	oFile.Close
:
:Loop Until offset >= Len(data)






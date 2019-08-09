setlocal
setlocal enabledelayedexpansion

if "%1"=="" goto :showUsage
if "%1"=="/?" goto :showUsage
if "%1"=="-?" goto :showUsage
if /I "%1"=="/HELP" goto :showUsage
goto :doIt

:showUsage
::===0=========1=========2=========3=========4=========5=========6=========7=========8
echo.
echo Signs an executable
echo.
echo Usage:
echo.
echo signmodule projectName [path] [/T:{DLL^|OCX^|EXE}] 
echo.
echo   projectName             Project name (excluding version).
echo.
echo   path                    Path to folder containing project file.
echo.
echo   /T                      Project type: DLL (default) or OCX
echo.
echo   On entry, ^%%BIN-PATH^%% must be set to the folder where the compiled object
echo   file is stored.
echo.
exit /B
::===0=========1=========2=========3=========4=========5=========6=========7=========8

:doIt
set PROJECTNAME=
set FOLDER=

:parse

if "%~1" == "" goto :parsingComplete

set ARG=%~1
if /I "%ARG%" == "/T:DLL" (
	set EXTENSION=dll
) else if /I "%ARG%" == "/T:OCX" (
	set EXTENSION=ocx
) else if /I "%ARG%" == "/T:EXE" (
	set EXTENSION=exe
) else if /I "%ARG:~0,3%" == "/T:" (
	set EXTENSION=
) else if "%ARG:~0,1%"=="/" (
	echo Invalid parameter '%ARG%'
	set ERROR=1
) else if not defined PROJECTNAME (
	set PROJECTNAME=%ARG%
) else if not defined FOLDER (
	set FOLDER=%ARG%
	pushd !FOLDER!
	if errorlevel 1 (
		echo Invalid folder parameter '!FOLDER!'
		set ERROR=1
	)
) else (
	echo Invalid parameter '%ARG%'
	set ERROR=1
)

shift
goto :parse
	
:parsingComplete

if not defined BIN-PATH (
	echo ^%%BIN-PATH^%% is not defined
	set ERROR=1
)
if not defined PROJECTNAME (
	echo Projectname parameter must be supplied
	set ERROR=1
)
if not defined EXTENSION (
	echo /T:{DLL^|OCX^|EXE} setting missing or invalid
	set ERROR=1
)
if defined ERROR goto :err

echo =================================
set FILENAME=%PROJECTNAME%%VB6-BUILD-MAJOR%%VB6-BUILD-MINOR%.%EXTENSION%
if defined FOLDER (
	set FILENAME=%FOLDER%\%FILENAME%
)
SET filename=%BIN-PATH%\%FILENAME%
echo Signing %FILENAME%

signtool sign /t http://timestamp.comodoca.com/rfc3161 %FILENAME%
if errorlevel 1 (
	echo Error signing file
	set ERROR=1
	goto :err
)

if defined FOLDER popd %FOLDER%
exit /B 0

:err
if defined FOLDER popd %FOLDER%
exit /B 1



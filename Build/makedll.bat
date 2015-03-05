@echo off

:: makedll.bat
::
:: builds a VB6 dll or ocx project
::
:: Parameters:
::   %1 Project name (excluding version)
::   %2 File extension ('dll' or 'ocx')
::   %3 Binary compatibility ('P' or 'B')
::   %4 'compat' if compatibility location is not the Bin folder

echo =================================
echo Building %1
set filenamestub=%1

set extension="dll"
if "%2" == "dll" set extension="dll"
if "%2" == "ocx" set extension="ocx"

set binarycompat="B"
if "%3" == "P" set binarycompat="P"
if "%3" == "B" set binarycompat="B"

set compat="no"
if "%4" == "compat" set compat="yes"

if exist %1\Prev goto directoryExists

echo Making %1\Prev directory
mkdir %1\Prev

:directoryExists
echo Copying previous binary
copy %BIN-PATH%\%filenamestub%%twutils-version%.%extension% %1\Prev\* 

echo Setting project compatibility
setprojectcomp %1\%1.vbp -mode:%binarycompat%
if errorlevel 1 pause

echo Compiling
vb6 /m %1\%1.vbp
if errorlevel 1 pause

echo Setting binary compatibility
setprojectcomp %1\%1.vbp -mode:B
if errorlevel 1 pause

if %compat% == "no" goto end
if exist %1\Compat goto compatDirectoryExists

echo Making %1\Compat directory
mkdir %1\Compat

:compatDirectoryExists
if %binarycompat% == "B" goto end
echo Copying binary to %1\Compat
copy %BIN-PATH%\%filenamestub%%twutils-version%.%extension% %1\Compat\* 

:end

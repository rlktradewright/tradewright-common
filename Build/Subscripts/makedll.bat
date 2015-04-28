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

call setVersion

set EXTENSION=dll
if "%2" == "dll" set EXTENSION=dll
if "%2" == "ocx" set EXTENSION=ocx

set BINARY_COMPAT=B
if "%3" == "P" set BINARY_COMPAT=P
if "%3" == "B" set BINARY_COMPAT=B

set COMPAT=no
if "%4" == "COMPAT" set COMPAT=yes
if "%4" == "compat" set COMPAT=yes

set FILENAME=%1%TWUTILS-MAJOR%%TWUTILS-MINOR%

if not exist %1\Prev (
	echo Making %1\Prev directory
	mkdir %1\Prev 
)

if exist %BIN-PATH%\%FILENAME%.%EXTENSION% (
	echo Copying previous binary
	copy %BIN-PATH%\%FILENAME%.%EXTENSION% %1\Prev\* 
)

echo Setting binary compatibility mode = %BINARY_COMPAT%; version = %TWUTILS-MAJOR%.%TWUTILS-MINOR%.%TWUTILS-REVISION%
echo ... for file: %1\%1.vbp 
setprojectcomp %1\%1.vbp %TWUTILS-REVISION% -mode:%BINARY_COMPAT%
if errorlevel 1 pause

echo Compiling
vb6 /m %1\%1.vbp
if errorlevel 1 pause

if exist %BIN-PATH%\%FILENAME%.lib (
	echo Deleting .lib file
	del %BIN-PATH%\%FILENAME%.lib 
)
if exist %BIN-PATH%\%FILENAME%.exp (
	echo Deleting .exp file
	del %BIN-PATH%\%FILENAME%.exp 
)

echo Setting binary compatibility mode = B
setprojectcomp %1\%1.vbp %TWUTILS-REVISION% -mode:B
if errorlevel 1 goto :err

if "%COMPAT%" == "yes" (
	if not exist %1\Compat (
		echo Making %1\Compat directory
		mkdir %1\Compat
		if errorlevel 1 goto :err
	)
	if not "%BINARY_COMPAT%" == "B" (
		echo Copying binary to %1\Compat
		copy %BIN-PATH%\%FILENAME%.%EXTENSION% %1\COMPAT\* 
		if errorlevel 1 goto :err
	)
)

generateAssemblyManifest %1 %2 EMBED
if errorlevel 1 goto :err

exit /B 0

:err
pause
exit /B 1


@echo off

:: generateAssemblyManifest.bat
::
:: builds the manifest for a VB6 exe, dll or ocx project
::
:: Parameters:
::   %1 Project name (excluding version)
::   %2 File extension ('dll' or 'ocx' or 'exe')
::   %3 Embed manifest as resource ('EMBED' or 'NOEMBED')
::   %4	Use Version 6 Common Controls ('V6CC' or 'NOV6CC')
echo Generating manifest for %1

call setVersion

set EXTENSION=.dll
if "%2" == "dll" set EXTENSION=.dll
if "%2" == "ocx" set EXTENSION=.ocx
if "%2" == "exe" set EXTENSION=.exe

set EMBED=NOEMBED
if "%3" == "EMBED" set EMBED=EMBED
if "%3" == "E" set EMBED=EMBED
if "%3" == "embed" set EMBED=EMBED
if "%3" == "e" set EMBED=EMBED
if "%3" == "NOEMBED" set EMBED=NOEMBED
if "%3" == "N" set EMBED=NOEMBED
if "%3" == "noembed" set EMBED=NOEMBED
if "%3" == "n" set EMBED=NOEMBED

set V6CC=NOV6CC
if "%4" == "V6CC" set V6CC=V6CC
if "%4" == "NOV6CC" set V6CC=NOV6CC
if "%4" == "v6cc" set V6CC=V6CC
if "%4" == "nov6cc" set V6CC=NOV6CC

set OBJECTFILENAME=%1%TWUTILS-MAJOR%%TWUTILS-MINOR%%EXTENSION%
if "%EXTENSION%"==".exe" (
        set MANIFESTFILENAME=%1%TWUTILS-MAJOR%%TWUTILS-MINOR%.exe.manifest
) else (
        set MANIFESTFILENAME=%1%TWUTILS-MAJOR%%TWUTILS-MINOR%.manifest
)

if "%V6CC%" == "V6CC" (
	GenerateManifest /Proj:%1\%1.vbp /Out:%BIN-PATH%\%MANIFESTFILENAME% /V6CC
) else (
	GenerateManifest /Proj:%1\%1.vbp /Out:%BIN-PATH%\%MANIFESTFILENAME%
)
if errorlevel 1 goto :err

:: ensure mt.exe can find object files when hashing
pushd %BIN-PATH%

echo Updating manifest hash
mt.exe -manifest %BIN-PATH%\%MANIFESTFILENAME% -hashupdate -nologo
if errorlevel 1 (
	popd		
	goto :err
)

popd

if "%3" == "EMBED" (
	if "%EXTENSION%"==".ocx" (
		echo Can't embed manifest as resource for .ocx
	) else (
		echo Embedding manifest as a resource
		mt.exe -manifest %BIN-PATH%\%MANIFESTFILENAME% -outputresource:%BIN-PATH%\%OBJECTFILENAME%;#1 -nologo
		if errorlevel 1 goto :err
		echo Deleting manifest file
		del %BIN-PATH%\%MANIFESTFILENAME%
	)
)

exit /B 0

:err
exit /B 1



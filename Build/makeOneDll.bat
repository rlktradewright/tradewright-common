@echo off
setlocal 

if "%1"=="" goto :showUsage
if "%1"=="/?" goto :showUsage
if "%1"=="-?" goto :showUsage
if /I "%1"=="/HELP" goto :showUsage
goto :doIt

:showUsage
::===0=========1=========2=========3=========4=========5=========6=========7=========8
echo Build a dll
echo .
echo Usage:
echo.
echo makeOnedll projectName [/T:{DLL^|OCX}] [/B:{P^|B^|N}] [/M:{N^|E^|F}] [/C]
echo.
echo   projectName      Project name (excluding version)
echo   /T               Project type: DLL (default) or OCX
echo   /B               Binary compatibility: 
echo                        B =^> binary compatibility (default)
echo                        P =^> project compatibility
echo                        N =^> no compatibility
echo   /M               Manifest requirement:
echo                        N =^> no manifest (default)
echo                        E =^> embed manifest in object file
echo                        F =^> freestanding manifest file
echo   /C               Indicates the compatibility location is the
echo                    project's Compat subfolder rather than the Bin 
echo                    folder
::===0=========1=========2=========3=========4=========5=========6=========7=========8
exit /B

:doIt

%TW-PROJECTS-DRIVE%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN-PATH=%TW-PROJECTS-PATH%\Bin\TradeWright.TradeWrightCommon

call setTradeWrightCommonVersion.bat

set BINARY_COMPAT=B

pushd %TW-PROJECTS-PATH%\src

call makedll.bat %~1 %1 %~2 %~3 %~4 %~5

popd


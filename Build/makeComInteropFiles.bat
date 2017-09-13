::=============================================================================+
::                                                                             +
::   This command file generates COM interop DLLs to enable the TradeWright    +
::   Common components to be used in .Net programs.                            +
::                                                                             +
::   Note that these interop DLLs are included in the TradeWright Common       +
::   install, so you should not need to run this file in normal circumstances. +
::                                                                             +
::   Before running this file, the TradeWright Common components must be       +
::   registered using the registerDlls.bat command file. If you need to use    +
::   any of the TradeWright Common ActiveX controls in your .Net program, they +
::   will need to remain registered to be used with the forms designer.        +
::                                                                             +
::   If you don't need to use the TradeWright Common ActiveX controls in your  +
::   .Net programs, and if you use registration-free COM to access the         +
::   TradeWright Common .dlls, then you can un-register all the TradeWright    +
::   Common files after running this command file.                             +
::                                                                             +
::   You should run this file from the Visual Studio Developer Command Prompt  +
::   because it uses the tlbimp.exe and aximp.exe programs which are already   +
::   in the Developer Command Prompt's path.                                   +
::                                                                             +
::=============================================================================+

@echo off
setlocal

echo =================================
echo Generating COM interop files
echo.

set BUILD=%CD%\Build
set TW=%CD%\Bin
if not exist "%TW%\TradeWright.Common" (
	echo You are not currently in the correct folder.
	echo.
	echo You must run this command from the folder above the Bin folder
	echo containing the TradeWright Common executables.
	goto :Err
)

set COMINTEROP=%TW%\TradeWright.Common.ComInterop
set TWUTILITIES=%TW%\TradeWright.Common
set TWWIN32API=%TWUTILITIES%\twwin32api.tlb

if exist %COMINTEROP% (
	del %COMINTEROP%\*.dll
) else (
	mkdir %COMINTEROP%
)

cd %COMINTEROP%

set SOURCE=%TW%\TradeWright.Common.ExternalComponents

set ASM_VERSION=

call :TlbImp TlbInf32

set SOURCE=%TWUTILITIES%

call %BUILD%\Setup\DeploymentScripts\setTradeWrightCommonVersion.bat
set ASM_VERSION=/asmversion:%VB6-BUILD-MAJOR%.%VB6-BUILD-MINOR%.%VB6-BUILD-REVISION%

call :TlbImp TWUtilities40
call :TlbImp ExtProps40
call :TlbImp ExtEvents40
call :TlbImp BusObjUtils40

call :AxImp TWControls40

call :TlbImp GraphicsUtils40
call :TlbImp LayeredGraphics40
call :TlbImp GraphObjUtils40
call :TlbImp GraphObj40
call :TlbImp SpriteControlLib40

exit /B 0

:Err
exit /B 1

:TlbImp
echo =================================
tlbimp "%SOURCE%\%1.dll" /out:%1.dll /tlbreference:"%TWWIN32API%" /namespace:%1 /silence:3008 /silence:3011 /silence:3012 %ASM_VERSION% %REFERENCE%
if errorlevel 1 goto :Err
set REFERENCE=%REFERENCE% /reference:%1.dll
echo.
goto :EOF

:TlbImpAx
tlbimp "%SOURCE%\%1.ocx" /out:%1.dll /tlbreference:"%TWWIN32API%" /namespace:%1 /silence:3008 /silence:3011 /silence:3012 %ASM_VERSION% %REFERENCE%
if errorlevel 1 goto :Err
set REFERENCE=%REFERENCE% /reference:%1.dll
goto :EOF

:AxImp
echo =================================
call :TlbImpAx %1
aximp "%SOURCE%\%1.ocx" /out:Ax%1.dll /rcw:%1.dll
if errorlevel 1 goto :Err
echo.
goto :EOF

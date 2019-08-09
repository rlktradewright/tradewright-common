@echo off
setlocal

%TW-PROJECTS-DRIVE%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN_PATH_ROOT=%TW-PROJECTS-PATH%\Bin

call setMyVersion.bat

echo =================================
echo Signing components for TradeWright.Common
echo.

set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.Common

call signModule.bat BusObjUtils /T:DLL
if errorlevel 1 pause
call signModule.bat ExtEvents /T:DLL
if errorlevel 1 pause
call signModule.bat ExtProps /T:DLL
if errorlevel 1 pause
call signModule.bat GraphicsUtils /T:DLL
if errorlevel 1 pause
call signModule.bat GraphObj /T:DLL
if errorlevel 1 pause
call signModule.bat GraphObjUtils /T:DLL
if errorlevel 1 pause
call signModule.bat LayeredGraphics /T:DLL
if errorlevel 1 pause
call signModule.bat SpriteControlLib /T:DLL
if errorlevel 1 pause
call signModule.bat TWControls /T:OCX
if errorlevel 1 pause
call signModule.bat TWUtilities /T:DLL
if errorlevel 1 pause



echo =================================
echo Signing deliverable projects
echo.

set BIN-PATH=%BIN_PATH_ROOT%

call signModule.bat ClockTester /T:EXE
if errorlevel 1 pause

call signModule.bat IntervalTimerTester /T:EXE
if errorlevel 1 pause

call signModule.bat LayeredGraphicsTest /T:EXE
if errorlevel 1 pause

call signModule.bat TasksDemo /T:EXE
if errorlevel 1 pause


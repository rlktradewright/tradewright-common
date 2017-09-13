@echo off
setlocal

%TW-PROJECTS-DRIVE%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN-PATH=%TW-PROJECTS-PATH%\Bin

call setMyVersion.bat

set DEP=/DEP:%TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build\ExternalDependencies.txt

echo =================================
echo Making sample apps
echo.

pushd %TW-PROJECTS-PATH%\src\SampleApps

call makeExe.bat ClockTester ClockTester /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat IntervalTimerTester IntervalTimerTester /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat LayeredGraphicsTest LayeredGraphicsTest /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat TasksDemo TasksDemo /M:E %DEP%
if errorlevel 1 pause

popd



@echo off
setlocal 

%TW-PROJECTS-DRIVE%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN-PATH=%TW-PROJECTS-PATH%\Bin

call setMyVersion.bat

pushd %TW-PROJECTS-PATH%\SampleApps

call makeExe.bat ClockTester ClockTester
call makeExe.bat IntervalTimerTester IntervalTimerTester 
call makeExe.bat LayeredGraphicsTest LayeredGraphicsTest 
call makeExe.bat TasksDemo TasksDemo 

popd
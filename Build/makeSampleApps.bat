@echo off

%TW-PROJECTS-DRIVE%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build;%PATH%

set BIN-PATH=%TW-PROJECTS-PATH%\Bin

pushd %TW-PROJECTS-PATH%\SampleApps

call makeExe ClockTester
call makeExe IntervalTimerTester
call makeExe LayeredGraphicsTest
call makeExe TasksDemo

popd
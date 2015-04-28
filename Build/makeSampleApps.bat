@echo off

%TW-PROJECTS-DRIVE%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build;%PATH%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build\Subscripts;%PATH%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build\Tools;%PATH%
path C:\Program Files\Microsoft SDKs\Windows\v6.0A\bin;%PATH%

set BIN-PATH=%TW-PROJECTS-PATH%\Bin

pushd %TW-PROJECTS-PATH%\SampleApps

call makeExe ClockTester V6CC
call makeExe IntervalTimerTester V6CC
call makeExe LayeredGraphicsTest V6CC
call makeExe TasksDemo V6CC

popd